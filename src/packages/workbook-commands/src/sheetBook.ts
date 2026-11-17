import { columnIndex, sheetOf, type SnapshotBook } from "./cells";

// Reads a saved Univer workbook the way the VBA reads Worksheet.Cells(r, c): rows and
// columns are 1-based. Writes are recorded in order and laid over later reads, so a port
// can run its VBA steps in sequence and hand the whole batch to the host at the end.

export interface CellLook {
  bold?: boolean;
  fontFamily?: string;
  fontSize?: number;
  hAlign?: "left" | "center" | "right";
  vAlign?: "middle";
}

export type CellScalar = string | number | boolean | null;

export type CellWrite =
  | { kind: "value"; sheet: string; row: number; col: number; value: CellScalar }
  /** Range.Value = data(): one block from `row`/`col`, rows may be ragged. */
  | { kind: "table"; sheet: string; row: number; col: number; values: CellScalar[][] }
  | { kind: "formula"; sheet: string; row: number; col: number; formula: string }
  /** `formats` is Range.Clear (contents and formatting); otherwise ClearContents. */
  | { kind: "clear"; sheet: string; row: number; col: number; row2: number; col2: number; formats?: boolean }
  | { kind: "fill"; sheet: string; row: number; col: number; row2: number; col2: number; fill: string | null }
  | { kind: "font"; sheet: string; row: number; col: number; row2: number; col2: number; color: string }
  | { kind: "format"; sheet: string; row: number; col: number; pattern: string }
  /** EnsureHiddenSheet: an existing sheet loses everything (Cells.Clear); a missing one is added hidden. */
  | { kind: "sheet"; sheet: string }
  | { kind: "merge"; sheet: string; row: number; col: number; row2: number; col2: number }
  | { kind: "unmerge"; sheet: string; row: number; col: number; row2: number; col2: number }
  | { kind: "look"; sheet: string; row: number; col: number; row2: number; col2: number; look: CellLook }
  | { kind: "border"; sheet: string; row: number; col: number; row2: number; col2: number; edge: "bottom" | "outside"; line: "thin" | "hair"; color: string };

export interface CellArea {
  row: number;
  col: number;
  row2: number;
  col2: number;
}

interface RawCell {
  v?: unknown;
  f?: string | null;
  si?: string | null;
  s?: unknown;
}

interface RawStyle {
  bg?: { rgb?: string | null } | null;
  cl?: { rgb?: string | null } | null;
  n?: { pattern?: string } | null;
  ff?: string;
  fs?: number;
  bl?: number;
}

interface RawSheet {
  id?: string;
  name?: string;
  cellData?: Record<string, Record<string, RawCell | null> | null> | null;
  mergeData?: Array<{ startRow: number; endRow: number; startColumn: number; endColumn: number }>;
  columnData?: Record<string, { w?: number } | null>;
  rowData?: Record<string, { h?: number } | null>;
  defaultColumnWidth?: number;
  defaultRowHeight?: number;
}

interface RawBook extends SnapshotBook {
  styles?: Record<string, RawStyle | null>;
  resources?: Array<{ name: string; data: string }>;
}

export interface ProtectionInfo {
  sheet: string;
  unlockedRanges: string[];
}

const ERROR_TEXT = /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|SPILL!|CALC!|GETTING_DATA)$/;

export function cellLetters(col: number): string {
  let out = "";
  let n = col;
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

export function a1Of(row: number, col: number): string {
  return `${cellLetters(col)}${row}`;
}

/** "B3:D9" or "B3" as 1-based bounds. */
export function parseArea(ref: string): CellArea | null {
  const match = /^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/i.exec(ref.trim());
  if (!match?.[1] || !match[2]) return null;
  const col = columnIndex(match[1]) + 1;
  const row = Number(match[2]);
  const col2 = match[3] ? columnIndex(match[3]) + 1 : col;
  const row2 = match[4] ? Number(match[4]) : row;
  return { row: Math.min(row, row2), col: Math.min(col, col2), row2: Math.max(row, row2), col2: Math.max(col, col2) };
}

export function isErrorValue(value: unknown): boolean {
  return typeof value === "string" && ERROR_TEXT.test(value.trim());
}

/** VBA IsNumberValue: a real number. Empty and "" (formula blank) do not count. */
export function isNumber(value: unknown): value is number {
  return typeof value === "number" && Number.isFinite(value);
}

/** VBA IsNumeric on a non-empty cell value, so numeric text counts. */
export function numeric(value: unknown): number | null {
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  if (typeof value === "boolean") return value ? -1 : 0;
  if (typeof value !== "string" || isErrorValue(value)) return null;
  const text = value.trim().replace(/,/g, "");
  if (text === "" || !/^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(text)) return null;
  return Number(text);
}

export function cellText(value: unknown): string {
  if (value == null || isErrorValue(value)) return "";
  return String(value).trim();
}

/** VBA Round / CLng: halves go to the even neighbour. */
export function bankersRound(value: number, digits = 0): number {
  const factor = 10 ** digits;
  const scaled = value * factor;
  const floor = Math.floor(scaled);
  const diff = scaled - floor;
  let rounded: number;
  if (Math.abs(diff - 0.5) < 1e-9) rounded = floor % 2 === 0 ? floor : floor + 1;
  else rounded = Math.round(scaled);
  return rounded / factor;
}

/** WorksheetFunction.Round and Format$: halves go away from zero. */
export function excelRound(value: number, digits = 0): number {
  const factor = 10 ** digits;
  const sign = value < 0 ? -1 : 1;
  return (sign * Math.round(Math.abs(value) * factor + 1e-9)) / factor;
}

/** Format$(x, "0.00"). */
export function fixed2(value: number): string {
  return excelRound(value, 2).toFixed(2);
}

function normalizeColor(rgb: string | null | undefined): string | null {
  if (!rgb) return null;
  const text = rgb.trim().replace(/^#/, "").toUpperCase();
  if (/^[0-9A-F]{8}$/.test(text)) return `#${text.slice(2)}`;
  if (/^[0-9A-F]{6}$/.test(text)) return `#${text}`;
  const match = /^RGBA?\((\d+),\s*(\d+),\s*(\d+)/i.exec(rgb.trim());
  if (match) return `#${[match[1], match[2], match[3]].map((part) => Number(part).toString(16).padStart(2, "0").toUpperCase()).join("")}`;
  return null;
}

export function rgb(r: number, g: number, b: number): string {
  return `#${[r, g, b].map((part) => part.toString(16).padStart(2, "0").toUpperCase()).join("")}`;
}

function key(sheet: string, row: number, col: number): string {
  return `${sheet}\u0000${row}\u0000${col}`;
}

function inside(area: CellArea, row: number, col: number): boolean {
  return row >= area.row && row <= area.row2 && col >= area.col && col <= area.col2;
}

export class SheetBook {
  readonly writes: CellWrite[] = [];
  private readonly values = new Map<string, unknown>();
  private readonly formulas = new Map<string, string | null>();
  private readonly fills = new Map<string, string | null>();
  private readonly fillAreas: Array<{ sheet: string; area: CellArea; fill: string | null }> = [];
  private readonly clears: Array<{ sheet: string; area: CellArea }> = [];
  private readonly unlocked = new Map<string, CellArea[]>();
  private readonly added = new Set<string>();

  constructor(readonly book: SnapshotBook, protection: ProtectionInfo[] = []) {
    for (const record of protection) {
      const areas = record.unlockedRanges.flatMap((ref) => {
        const area = parseArea(ref);
        return area ? [area] : [];
      });
      this.unlocked.set(record.sheet, areas);
    }
  }

  hasSheet(name: string): boolean {
    return this.added.has(name) || sheetOf(this.book, name) != null;
  }

  value(sheet: string, row: number, col: number): unknown {
    const id = key(sheet, row, col);
    if (this.values.has(id)) return this.values.get(id);
    for (let i = this.clears.length - 1; i >= 0; i -= 1) {
      const clear = this.clears[i]!;
      if (clear.sheet === sheet && inside(clear.area, row, col)) return null;
    }
    return this.raw(sheet, row, col)?.v ?? null;
  }

  at(sheet: string, a1: string): unknown {
    const area = parseArea(a1);
    return area ? this.value(sheet, area.row, area.col) : null;
  }

  text(sheet: string, row: number, col: number): string {
    return cellText(this.value(sheet, row, col));
  }

  hasFormula(sheet: string, row: number, col: number): boolean {
    const id = key(sheet, row, col);
    if (this.formulas.has(id)) return this.formulas.get(id) != null;
    if (this.values.has(id)) return false;
    for (const clear of this.clears) if (clear.sheet === sheet && inside(clear.area, row, col)) return false;
    const cell = this.raw(sheet, row, col);
    return Boolean(cell?.f) || Boolean(cell?.si);
  }

  formula(sheet: string, row: number, col: number): string | null {
    const id = key(sheet, row, col);
    if (this.formulas.has(id)) return this.formulas.get(id) ?? null;
    if (!this.hasFormula(sheet, row, col)) return null;
    const f = this.raw(sheet, row, col)?.f;
    return typeof f === "string" && f.length > 0 ? f : null;
  }

  /** Static fill (Interior.Color) as "#RRGGBB", or null for no fill. */
  fill(sheet: string, row: number, col: number): string | null {
    const id = key(sheet, row, col);
    if (this.fills.has(id)) return this.fills.get(id) ?? null;
    for (let i = this.fillAreas.length - 1; i >= 0; i -= 1) {
      const paint = this.fillAreas[i]!;
      if (paint.sheet === sheet && inside(paint.area, row, col)) return paint.fill;
    }
    return normalizeColor(this.style(sheet, row, col)?.bg?.rgb ?? null);
  }

  style(sheet: string, row: number, col: number): RawStyle | null {
    const s = this.raw(sheet, row, col)?.s;
    if (s == null) return null;
    if (typeof s === "string") return (this.book as RawBook).styles?.[s] ?? null;
    return s as RawStyle;
  }

  locked(sheet: string, row: number, col: number): boolean {
    const areas = this.unlocked.get(sheet);
    if (!areas) return false;
    return !areas.some((area) => inside(area, row, col));
  }

  /** Last row at or above `fromRow` with anything in the column (End(xlUp)). */
  lastUsedRow(sheet: string, col: number, fromRow = 1_048_576): number {
    const rows = this.rows(sheet);
    let best = 1;
    for (const rowKey of Object.keys(rows ?? {})) {
      const row = Number(rowKey) + 1;
      if (row > fromRow || row <= best) continue;
      if (cellText(this.value(sheet, row, col)) !== "" || isNumber(this.value(sheet, row, col))) best = row;
    }
    for (const write of this.writes) {
      if (write.kind === "value" && write.sheet === sheet && write.col === col && write.row <= fromRow && write.row > best && write.value != null && write.value !== "") best = write.row;
      if (write.kind === "table" && write.sheet === sheet && col >= write.col) {
        for (let i = write.values.length - 1; i >= 0; i -= 1) {
          const row = write.row + i;
          if (row > fromRow) continue;
          if (row <= best) break;
          const value = write.values[i]?.[col - write.col];
          if (value != null && value !== "" && this.value(sheet, row, col) === value) {
            best = row;
            break;
          }
        }
      }
    }
    return best;
  }

  columnWidth(sheet: string, col: number): number {
    const raw = sheetOf(this.book, sheet) as RawSheet | null;
    return raw?.columnData?.[col - 1]?.w ?? raw?.defaultColumnWidth ?? 73;
  }

  rowHeight(sheet: string, row: number): number {
    const raw = sheetOf(this.book, sheet) as RawSheet | null;
    return raw?.rowData?.[row - 1]?.h ?? raw?.defaultRowHeight ?? 19;
  }

  mergeAt(sheet: string, row: number, col: number): CellArea {
    const raw = sheetOf(this.book, sheet) as RawSheet | null;
    for (const merge of raw?.mergeData ?? []) {
      const area = { row: merge.startRow + 1, col: merge.startColumn + 1, row2: merge.endRow + 1, col2: merge.endColumn + 1 };
      if (inside(area, row, col)) return area;
    }
    return { row, col, row2: row, col2: col };
  }

  /** Items of the cell's list validation (range or comma list), blanks skipped. */
  validationList(sheet: string, row: number, col: number): string[] {
    const raw = sheetOf(this.book, sheet) as RawSheet | null;
    const sheetId = raw?.id ?? Object.entries(this.book.sheets ?? {}).find(([, value]) => value?.name === sheet)?.[0];
    const resource = (this.book as RawBook).resources?.find((entry) => entry.name === "SHEET_DATA_VALIDATION_PLUGIN");
    if (!sheetId || !resource) return [];
    let rules: Array<{ type?: string; ranges?: Array<{ startRow: number; endRow: number; startColumn: number; endColumn: number }>; formula1?: string }> = [];
    try {
      const parsed = JSON.parse(resource.data) as Record<string, typeof rules>;
      rules = parsed[sheetId] ?? [];
    } catch {
      return [];
    }
    const rule = rules.find((entry) => entry.type === "list" && entry.ranges?.some((range) => row - 1 >= range.startRow && row - 1 <= range.endRow && col - 1 >= range.startColumn && col - 1 <= range.endColumn));
    const formula = rule?.formula1?.trim() ?? "";
    if (!formula) return [];
    if (formula.startsWith("=")) {
      const area = parseArea(formula.slice(1).replace(/^'?[^!]*'?!/, ""));
      if (!area) return [];
      const items: string[] = [];
      for (let r = area.row; r <= area.row2; r += 1) {
        for (let c = area.col; c <= area.col2; c += 1) {
          const text = this.text(sheet, r, c);
          if (text) items.push(text);
        }
      }
      return items;
    }
    if (formula.startsWith("[")) {
      try {
        return (JSON.parse(formula) as unknown[]).map((item) => String(item).trim()).filter(Boolean);
      } catch {
        return [];
      }
    }
    return formula.replace(/^"|"$/g, "").split(",").map((item) => item.trim()).filter(Boolean);
  }

  setValue(sheet: string, row: number, col: number, value: CellScalar): void {
    const id = key(sheet, row, col);
    this.values.set(id, value);
    this.formulas.set(id, null);
    this.writes.push({ kind: "value", sheet, row, col, value });
  }

  setTable(sheet: string, row: number, col: number, values: CellScalar[][]): void {
    values.forEach((line, i) => {
      line.forEach((value, j) => {
        const id = key(sheet, row + i, col + j);
        this.values.set(id, value);
        this.formulas.set(id, null);
      });
    });
    this.writes.push({ kind: "table", sheet, row, col, values });
  }

  setFormula(sheet: string, row: number, col: number, formula: string): void {
    const id = key(sheet, row, col);
    this.values.delete(id);
    this.formulas.set(id, formula);
    this.writes.push({ kind: "formula", sheet, row, col, formula });
  }

  clear(sheet: string, ref: string | CellArea, formats = false): void {
    const area = typeof ref === "string" ? parseArea(ref) : ref;
    if (!area) return;
    this.forget(sheet, area, formats);
    this.writes.push({ kind: "clear", sheet, ...area, ...(formats ? { formats } : {}) });
  }

  /** EnsureHiddenSheet / EnsureHidden: blank the sheet, adding it hidden when the workbook has none. */
  resetSheet(sheet: string): void {
    const area = { row: 1, col: 1, row2: 1_048_576, col2: 16_384 };
    this.forget(sheet, area, true);
    this.added.add(sheet);
    this.writes.push({ kind: "sheet", sheet });
  }

  merge(sheet: string, ref: string | CellArea): void {
    const area = typeof ref === "string" ? parseArea(ref) : ref;
    if (area && (area.row !== area.row2 || area.col !== area.col2)) this.writes.push({ kind: "merge", sheet, ...area });
  }

  unmerge(sheet: string, ref: string | CellArea): void {
    const area = typeof ref === "string" ? parseArea(ref) : ref;
    if (area) this.writes.push({ kind: "unmerge", sheet, ...area });
  }

  setLook(sheet: string, ref: string | CellArea, look: CellLook): void {
    const area = typeof ref === "string" ? parseArea(ref) : ref;
    if (area) this.writes.push({ kind: "look", sheet, ...area, look });
  }

  setBorder(sheet: string, ref: string | CellArea, edge: "bottom" | "outside", line: "thin" | "hair", color: string): void {
    const area = typeof ref === "string" ? parseArea(ref) : ref;
    if (area) this.writes.push({ kind: "border", sheet, ...area, edge, line, color });
  }

  private forget(sheet: string, area: CellArea, formats: boolean): void {
    const within = (id: string) => {
      const [s, r, c] = id.split("\u0000");
      return s === sheet && inside(area, Number(r), Number(c));
    };
    for (const id of [...this.values.keys()]) if (within(id)) this.values.delete(id);
    for (const id of [...this.formulas.keys()]) if (within(id)) this.formulas.set(id, null);
    this.clears.push({ sheet, area });
    if (formats) {
      for (const id of [...this.fills.keys()]) if (within(id)) this.fills.delete(id);
      this.fillAreas.push({ sheet, area, fill: null });
    }
  }

  setFill(sheet: string, ref: string | CellArea, fill: string | null): void {
    const area = typeof ref === "string" ? parseArea(ref) : ref;
    if (!area) return;
    if (area.row === area.row2 && area.col === area.col2) this.fills.set(key(sheet, area.row, area.col), fill);
    else {
      for (const id of [...this.fills.keys()]) {
        const [s, r, c] = id.split("\u0000");
        if (s === sheet && inside(area, Number(r), Number(c))) this.fills.delete(id);
      }
      this.fillAreas.push({ sheet, area, fill });
    }
    this.writes.push({ kind: "fill", sheet, ...area, fill });
  }

  setFontColor(sheet: string, ref: string | CellArea, color: string): void {
    const area = typeof ref === "string" ? parseArea(ref) : ref;
    if (!area) return;
    this.writes.push({ kind: "font", sheet, ...area, color });
  }

  setFormat(sheet: string, row: number, col: number, pattern: string): void {
    this.writes.push({ kind: "format", sheet, row, col, pattern });
  }

  private rows(sheet: string): Record<string, Record<string, RawCell | null> | null> | null {
    return ((sheetOf(this.book, sheet) as RawSheet | null)?.cellData as Record<string, Record<string, RawCell | null> | null> | undefined) ?? null;
  }

  private raw(sheet: string, row: number, col: number): RawCell | null {
    return this.rows(sheet)?.[row - 1]?.[col - 1] ?? null;
  }
}
