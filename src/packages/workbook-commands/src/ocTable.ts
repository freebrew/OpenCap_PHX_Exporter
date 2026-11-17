import { bankersRound, cellText, type SheetBook } from "./sheetBook";

// VBA value conversions the OpenCap ports share, and the _OC_* sheet reader.

/** VBA Trim$: spaces only, so a tab or line break stays part of the text. */
export function vbTrim(text: string): string {
  return text.replace(/^ +| +$/g, "");
}

/** VBA IsNumeric on text: signs, a currency $, thousands commas, decimals and an exponent. */
export function vbaIsNumeric(text: string): boolean {
  const t = text.trim();
  return /^[+-]?\$?(\d[\d,]*(\.\d*)?|\.\d+)(e[+-]?\d+)?$/i.test(t) || /^\(\$?\d[\d,]*(\.\d*)?\)$/.test(t);
}

export function vbaCdbl(text: string): number {
  const t = text.trim().replace(/[$,]/g, "");
  if (t.startsWith("(") && t.endsWith(")")) return -Number(t.slice(1, -1));
  return Number(t);
}

/** CStr of a cell value the way VBA prints a Double (15 significant digits). */
export function vbaStr(value: unknown): string {
  if (value == null) return "";
  if (typeof value === "number") {
    if (!Number.isFinite(value)) return "";
    if (Number.isInteger(value) && Math.abs(value) < 1e15) return String(value);
    return String(Number(value.toPrecision(15)));
  }
  if (typeof value === "boolean") return value ? "True" : "False";
  return String(value);
}

/** NumOrZero: IsNumeric then CDbl, else 0. */
export function numOrZero(value: unknown): number {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  if (typeof value === "boolean") return value ? -1 : 0;
  if (typeof value === "string" && vbaIsNumeric(value)) return vbaCdbl(value);
  return 0;
}

/** BhaNumber: IsNumeric then CLng, else 0. */
export function bhaNumber(value: unknown): number {
  const n = numOrZero(value);
  return n === 0 ? 0 : bankersRound(n);
}

/** One _OC_* sheet: row 1 headers, End(xlUp) per column, HeaderCol by trimmed name. */
export class OcTable {
  private readonly lastCol: number;

  constructor(private readonly book: SheetBook, readonly sheet: string) {
    let last = 0;
    if (book.hasSheet(sheet)) {
      for (let col = 1; col <= 512; col += 1) if (cellText(book.value(sheet, 1, col)) !== "") last = col;
    }
    this.lastCol = last;
  }

  get exists(): boolean {
    return this.book.hasSheet(this.sheet);
  }

  /** HeaderCol / ColByName: the first name that matches a trimmed header, case-insensitive; 0 when none. */
  col(...names: string[]): number {
    for (const name of names) {
      for (let col = 1; col <= this.lastCol; col += 1) {
        if (cellText(this.book.value(this.sheet, 1, col)).toLowerCase() === name.toLowerCase()) return col;
      }
    }
    return 0;
  }

  lastRow(col: number): number {
    return col > 0 && this.exists ? this.book.lastUsedRow(this.sheet, col) : 1;
  }

  value(row: number, col: number): unknown {
    return col > 0 ? this.book.value(this.sheet, row, col) : null;
  }

  /** CStr(cell.Value & "") */
  str(row: number, col: number): string {
    return col > 0 ? vbaStr(this.book.value(this.sheet, row, col)) : "";
  }
}
