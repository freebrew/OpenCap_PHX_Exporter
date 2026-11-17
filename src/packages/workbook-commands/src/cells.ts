export interface SnapshotSheet {
  name?: string;
  cellData?: Record<string, Record<string, { v?: unknown; f?: string } | null> | null>;
}

export interface SnapshotBook {
  sheets?: Record<string, SnapshotSheet>;
  sheetOrder?: string[];
}

export function columnIndex(letters: string): number {
  return letters.toUpperCase().split("").reduce((total, letter) => total * 26 + letter.charCodeAt(0) - 64, 0) - 1;
}

export function parseA1(a1: string): { row: number; column: number } | null {
  const match = /^([A-Z]+)(\d+)$/i.exec(a1.trim());
  if (!match?.[1] || !match[2]) return null;
  return { column: columnIndex(match[1]), row: Number(match[2]) - 1 };
}

export function sheetOf(book: SnapshotBook, name: string): SnapshotSheet | null {
  const sheets = book.sheets ?? {};
  return Object.values(sheets).find((sheet) => sheet?.name === name) ?? null;
}

export function readCell(book: SnapshotBook, sheetName: string, a1: string): unknown {
  const address = parseA1(a1);
  const sheet = sheetOf(book, sheetName);
  if (!address || !sheet?.cellData) return null;
  return sheet.cellData[address.row]?.[address.column]?.v ?? null;
}

export function readColumn(book: SnapshotBook, sheetName: string, column: string, fromRow: number, toRow: number): unknown[] {
  const sheet = sheetOf(book, sheetName);
  const index = columnIndex(column);
  const values: unknown[] = [];
  for (let row = fromRow; row <= toRow; row += 1) {
    values.push(sheet?.cellData?.[row - 1]?.[index]?.v ?? null);
  }
  return values;
}

export function readNumber(book: SnapshotBook, sheetName: string, a1: string): number | null {
  const value = readCell(book, sheetName, a1);
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}