import { extractFunctionCalls, extractRefs, sheetTitle, unknownFunctions, type CellFormulaRecord, type DefinedNameRecord } from "../../xlsx-bridge/src/index.ts";

export type FormulaClass = "evaluable" | "udf-dependent" | "mapper-bug";

const NAME_TOKEN = /[A-Za-z_][A-Za-z0-9_.]*/g;
const SMALL_AREA = 48;

interface Box {
  sheet: string;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

function identifiers(formula: string): string[] {
  const without = formula.replace(/"(?:[^"]|"")*"/g, "\"\"");
  return [...without.matchAll(NAME_TOKEN)].map((match) => match[0]);
}

function isDirectUdf(formula: string): boolean {
  return unknownFunctions(extractFunctionCalls(formula)).length > 0;
}

export function cellId(formula: CellFormulaRecord): string {
  return `${formula.sheet}\u0000${formula.row}\u0000${formula.col}`;
}

function pointId(sheet: string, row: number, col: number): string {
  return `${sheet}\u0000${row}\u0000${col}`;
}

export function classifyFormulas(formulas: CellFormulaRecord[], names: DefinedNameRecord[]): Map<string, FormulaClass> {
  const knownNames = new Set(names.map((name) => name.name.toUpperCase()));
  const udfNames = new Set<string>();
  const referencesUdfName = (formula: string): boolean => identifiers(formula).some((token) => udfNames.has(token.toUpperCase()) && knownNames.has(token.toUpperCase()));
  let grew = true;
  while (grew) {
    grew = false;
    for (const name of names) {
      const key = name.name.toUpperCase();
      if (udfNames.has(key)) continue;
      if (isDirectUdf(name.text) || referencesUdfName(name.text)) {
        udfNames.add(key);
        grew = true;
      }
    }
  }

  const classes = new Map<string, FormulaClass>();
  const boxes = new Map<string, Box[]>();
  const dependents = new Map<string, string[]>();
  const present = new Set<string>();
  for (const formula of formulas) {
    const id = cellId(formula);
    present.add(id);
    const usesUdfName = referencesUdfName(formula.formula);
    classes.set(id, isDirectUdf(formula.formula) || usesUdfName ? "udf-dependent" : "evaluable");
    const large: Box[] = [];
    for (const ref of extractRefs(formula.formula)) {
      if (ref.kind === "cols" || ref.kind === "rows") continue;
      const sheet = sheetTitle(ref.sheet) ?? formula.sheet;
      const r1 = Math.min(ref.row1, ref.row2);
      const r2 = Math.max(ref.row1, ref.row2);
      const c1 = Math.min(ref.col1, ref.col2);
      const c2 = Math.max(ref.col1, ref.col2);
      const area = (r2 - r1 + 1) * (c2 - c1 + 1);
      if (area > SMALL_AREA) {
        large.push({ sheet, r1, c1, r2, c2 });
        continue;
      }
      for (let row = r1; row <= r2; row += 1) {
        for (let col = c1; col <= c2; col += 1) {
          const source = pointId(sheet, row, col);
          if (source === id) continue;
          const list = dependents.get(source);
          if (list) list.push(id);
          else dependents.set(source, [id]);
        }
      }
    }
    if (large.length > 0) boxes.set(id, large);
  }

  const udfCells: CellFormulaRecord[] = [];
  const queue: string[] = [];
  for (const formula of formulas) {
    const id = cellId(formula);
    if (classes.get(id) === "udf-dependent") {
      udfCells.push(formula);
      queue.push(id);
    }
  }
  const spread = () => {
    while (queue.length > 0) {
      const id = queue.pop();
      if (!id) break;
      for (const dependent of dependents.get(id) ?? []) {
        if (!present.has(dependent) || classes.get(dependent) === "udf-dependent") continue;
        classes.set(dependent, "udf-dependent");
        queue.push(dependent);
      }
    }
  };
  spread();

  const udfBySheet = new Map<string, CellFormulaRecord[]>();
  const remember = (formula: CellFormulaRecord) => {
    const list = udfBySheet.get(formula.sheet);
    if (list) list.push(formula);
    else udfBySheet.set(formula.sheet, [formula]);
  };
  for (const formula of udfCells) remember(formula);

  let guard = 0;
  let marked = true;
  while (marked && guard < 64) {
    guard += 1;
    marked = false;
    for (const formula of formulas) {
      const id = cellId(formula);
      if (classes.get(id) === "udf-dependent") continue;
      const ranges = boxes.get(id);
      if (!ranges) continue;
      for (const box of ranges) {
        const candidates = udfBySheet.get(box.sheet) ?? [];
        for (const candidate of candidates) {
          if (candidate.row < box.r1 || candidate.row > box.r2 || candidate.col < box.c1 || candidate.col > box.c2) continue;
          classes.set(id, "udf-dependent");
          queue.push(id);
          remember(formula);
          marked = true;
          break;
        }
        if (classes.get(id) === "udf-dependent") break;
      }
    }
    spread();
  }
  return classes;
}
