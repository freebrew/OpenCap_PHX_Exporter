export interface CellPos {
  row: number;
  col: number;
}

export interface RefRange {
  sheet: string | null;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export function colToIndex(col: string): number {
  let n = 0;
  const letters = col.toUpperCase();
  for (let i = 0; i < letters.length; i += 1) {
    n = n * 26 + (letters.charCodeAt(i) - 64);
  }
  return n - 1;
}

export function indexToCol(index: number): string {
  let n = index + 1;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

export function parseA1(a1: string): CellPos {
  const match = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(a1.trim());
  if (!match) {
    throw new Error(`Not an A1 cell: ${a1}`);
  }
  return { row: Number(match[4]) - 1, col: colToIndex(match[2]) };
}

export function toA1(row: number, col: number): string {
  return `${indexToCol(col)}${row + 1}`;
}

export function toRangeA1(r1: number, c1: number, r2: number, c2: number): string {
  const start = toA1(Math.min(r1, r2), Math.min(c1, c2));
  const end = toA1(Math.max(r1, r2), Math.max(c1, c2));
  return start === end ? start : `${start}:${end}`;
}

const SHEET_PREFIX = /^(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!/;
const CELL = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)/;
const COL_RANGE = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9])/;
const ROW_RANGE = /^(\$?)(\d+):(\$?)(\d+)(?!\d)/;

function isNameChar(ch: string | undefined): boolean {
  return ch !== undefined && /[A-Za-z0-9_.]/.test(ch);
}

export interface RawRef {
  start: number;
  end: number;
  sheet: string;
  kind: "cell" | "range" | "cols" | "rows";
  absCol1: boolean;
  absRow1: boolean;
  absCol2: boolean;
  absRow2: boolean;
  col1: number;
  row1: number;
  col2: number;
  row2: number;
}

function sheetNameOf(prefix: string): string {
  if (!prefix) return "";
  const body = prefix.slice(0, -1);
  if (body.startsWith("'")) {
    return body.slice(1, -1).replace(/''/g, "'");
  }
  return body;
}

/** Read one A1 reference or range starting at `index`, or null. */
export function readRef(formula: string, index: number): RawRef | null {
  if (isNameChar(formula[index - 1])) return null;
  let cursor = index;
  let sheet = "";
  const sheetMatch = SHEET_PREFIX.exec(formula.slice(index));
  if (sheetMatch) {
    sheet = sheetMatch[0];
    cursor = index + sheet.length;
  }
  const rest = formula.slice(cursor);
  const cell = CELL.exec(rest);
  if (cell) {
    const after = rest.slice(cell[0].length);
    if (after.startsWith(":")) {
      const cell2 = CELL.exec(after.slice(1));
      if (cell2 && !after.slice(1 + cell2[0].length).startsWith("(")) {
        const end = cursor + cell[0].length + 1 + cell2[0].length;
        return {
          start: index,
          end,
          sheet,
          kind: "range",
          absCol1: cell[1] === "$",
          col1: colToIndex(cell[2]),
          absRow1: cell[3] === "$",
          row1: Number(cell[4]) - 1,
          absCol2: cell2[1] === "$",
          col2: colToIndex(cell2[2]),
          absRow2: cell2[3] === "$",
          row2: Number(cell2[4]) - 1,
        };
      }
    }
    if (!after.startsWith("(")) {
      return {
        start: index,
        end: cursor + cell[0].length,
        sheet,
        kind: "cell",
        absCol1: cell[1] === "$",
        col1: colToIndex(cell[2]),
        absRow1: cell[3] === "$",
        row1: Number(cell[4]) - 1,
        absCol2: cell[1] === "$",
        col2: colToIndex(cell[2]),
        absRow2: cell[3] === "$",
        row2: Number(cell[4]) - 1,
      };
    }
  }
  const cols = COL_RANGE.exec(rest);
  if (cols) {
    return {
      start: index,
      end: cursor + cols[0].length,
      sheet,
      kind: "cols",
      absCol1: cols[1] === "$",
      col1: colToIndex(cols[2]),
      absRow1: true,
      row1: 0,
      absCol2: cols[3] === "$",
      col2: colToIndex(cols[4]),
      absRow2: true,
      row2: 0,
    };
  }
  const rows = ROW_RANGE.exec(rest);
  if (rows && !sheet) {
    return {
      start: index,
      end: cursor + rows[0].length,
      sheet,
      kind: "rows",
      absCol1: true,
      col1: 0,
      absRow1: rows[1] === "$",
      row1: Number(rows[2]) - 1,
      absCol2: true,
      col2: 0,
      absRow2: rows[3] === "$",
      row2: Number(rows[4]) - 1,
    };
  }
  return null;
}

function formatCell(absCol: boolean, col: number, absRow: boolean, row: number): string {
  return `${absCol ? "$" : ""}${indexToCol(col)}${absRow ? "$" : ""}${row + 1}`;
}

function shiftIndex(value: number, delta: number, absolute: boolean): number {
  if (absolute) return value;
  return value + delta;
}

export function shiftFormula(formula: string, dRow: number, dCol: number): string {
  if (dRow === 0 && dCol === 0) return formula;
  let out = "";
  let i = 0;
  while (i < formula.length) {
    if (formula[i] === '"') {
      let j = i + 1;
      while (j < formula.length) {
        if (formula[j] === '"') {
          if (formula[j + 1] === '"') {
            j += 2;
            continue;
          }
          j += 1;
          break;
        }
        j += 1;
      }
      out += formula.slice(i, j);
      i = j;
      continue;
    }
    const ref = readRef(formula, i);
    if (!ref) {
      out += formula[i];
      i += 1;
      continue;
    }
    const col1 = shiftIndex(ref.col1, dCol, ref.absCol1);
    const row1 = shiftIndex(ref.row1, dRow, ref.absRow1);
    const col2 = shiftIndex(ref.col2, dCol, ref.absCol2);
    const row2 = shiftIndex(ref.row2, dRow, ref.absRow2);
    if (ref.kind === "cell") {
      out += `${ref.sheet}${formatCell(ref.absCol1, col1, ref.absRow1, row1)}`;
    } else if (ref.kind === "range") {
      out += `${ref.sheet}${formatCell(ref.absCol1, col1, ref.absRow1, row1)}:${formatCell(ref.absCol2, col2, ref.absRow2, row2)}`;
    } else if (ref.kind === "cols") {
      const a = `${ref.absCol1 ? "$" : ""}${indexToCol(col1)}`;
      const b = `${ref.absCol2 ? "$" : ""}${indexToCol(col2)}`;
      out += `${ref.sheet}${a}:${b}`;
    } else {
      const a = `${ref.absRow1 ? "$" : ""}${row1 + 1}`;
      const b = `${ref.absRow2 ? "$" : ""}${row2 + 1}`;
      out += `${a}:${b}`;
    }
    i = ref.end;
  }
  return out;
}

export function hasExternalRef(formula: string): boolean {
  return /\[[0-9]+\]/.test(stripStrings(formula));
}

export function stripStrings(formula: string): string {
  let out = "";
  let i = 0;
  while (i < formula.length) {
    if (formula[i] === '"') {
      out += '""';
      i += 1;
      while (i < formula.length) {
        if (formula[i] === '"') {
          if (formula[i + 1] === '"') {
            i += 2;
            continue;
          }
          i += 1;
          break;
        }
        i += 1;
      }
      continue;
    }
    out += formula[i];
    i += 1;
  }
  return out;
}

const CALL = /([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g;

export function extractFunctionCalls(formula: string): string[] {
  const names: string[] = [];
  const source = stripStrings(formula);
  for (const match of source.matchAll(CALL)) {
    const index = match.index ?? 0;
    if (isNameChar(source[index - 1])) continue;
    names.push(match[1].toUpperCase());
  }
  return names;
}

export function extractRefs(formula: string): RawRef[] {
  const source = stripStrings(formula);
  const refs: RawRef[] = [];
  let i = 0;
  while (i < source.length) {
    const ref = readRef(source, i);
    if (ref) {
      refs.push(ref);
      i = ref.end;
      continue;
    }
    i += 1;
  }
  return refs;
}

export function parseRangeRef(text: string): RefRange | null {
  const trimmed = text.trim().replace(/^=/, "");
  const ref = readRef(trimmed, 0);
  if (!ref || ref.end !== trimmed.length) return null;
  if (ref.kind === "cols" || ref.kind === "rows") return null;
  return {
    sheet: ref.sheet ? sheetNameOf(ref.sheet) : null,
    r1: Math.min(ref.row1, ref.row2),
    c1: Math.min(ref.col1, ref.col2),
    r2: Math.max(ref.row1, ref.row2),
    c2: Math.max(ref.col1, ref.col2),
  };
}

export function sheetTitle(prefix: string): string | null {
  if (!prefix) return null;
  return sheetNameOf(prefix);
}
