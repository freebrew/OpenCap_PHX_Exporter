// Module23's SelectBHA and the serial-hours table, read from the _OC_* sheets the OpenCap
// refresh already fills (Excel reads the same rows from its _FC_* copies). The dashboard's
// shapes and colours are the sheet's own; what comes across is the data those buttons show.

import { bhaNumber as bhaNumberOf, numOrZero, OcTable } from "./ocTable";
import { cellText, type CellScalar, type SheetBook } from "./sheetBook";

const BHA_SHEET = "_OC_BHA";
const INV_SHEET = "_OC_Inventory";
const DATA = "Data";
const SERIAL_HEADER_ROW = 2;
const SERIAL_FIRST_ROW = 3;
const SERIAL_CLEAR_TO = 400;

export interface BhaComponent {
  index: number;
  serial: string;
  itemCode: string;
  description: string;
  hrsSlid: number;
  hrsRot: number;
  hrsTotal: number;
  bha: number;
}

export interface BhaSummary {
  bha: number;
  section: string;
  status: string;
  motor: string;
  guidance: string;
  metres: number;
  hrsTotal: number;
  hrsSlid: number;
  hrsRot: number;
  hrsCirc: number;
  hrsBelowRot: number;
  activatedOn: string;
  completedOn: string;
  /** BHA Hrs Slid / BHA Total Hrs, 0 when the BHA has no hours. */
  slideFraction: number;
  components: BhaComponent[];
}

export interface SerialHoursRow {
  serial: string;
  description: string;
  circHours: number;
  /** Hours since last service. Past 300 the VBA paints the cell red. */
  hsls: number;
  bha: number;
}

type Row = Map<string, CellScalar>;

function rowsOf(book: SheetBook, sheet: string): Row[] {
  const table = new OcTable(book, sheet);
  if (!table.exists) return [];
  const headers: Array<[number, string]> = [];
  for (let col = 1; col <= 512; col += 1) {
    const header = cellText(book.value(sheet, 1, col)).trim();
    if (header) headers.push([col, header]);
  }
  const out: Row[] = [];
  const anchor = headers[0]?.[0] ?? 1;
  const limit = Math.min(book.lastUsedRow(sheet, anchor), 5000);
  for (let row = 2; row <= limit; row += 1) {
    const record: Row = new Map();
    for (const [col, header] of headers) record.set(header, book.value(sheet, row, col) as CellScalar);
    out.push(record);
  }
  return out;
}

function text(row: Row, header: string): string {
  return cellText(row.get(header) ?? null).trim();
}

function num(row: Row, header: string): number {
  return numOrZero(row.get(header) ?? null);
}

function bhaNumber(row: Row): number | null {
  const n = bhaNumberOf(row.get("BHA #") ?? null);
  return n > 0 ? n : null;
}

/** The BHA numbers the sheet holds, ascending, so the buttons can offer exactly those. */
export function bhaNumbers(book: SheetBook): number[] {
  const seen = new Set<number>();
  for (const row of rowsOf(book, BHA_SHEET)) {
    const n = bhaNumber(row);
    if (n != null) seen.add(n);
  }
  return [...seen].sort((a, b) => a - b);
}

/** SelectBHA's detail: the BHA's own metrics plus one row per component that has a serial, code or description. */
export function bhaSummary(book: SheetBook, bha: number): BhaSummary | null {
  const rows = rowsOf(book, BHA_SHEET).filter((row) => bhaNumber(row) === bha);
  const first = rows[0];
  if (!first) return null;
  const metres = rows.map((row) => num(row, "Metres Drilled")).find((value) => value !== 0) ?? 0;
  const hrsTotal = num(first, "BHA Total Hrs");
  const hrsSlid = num(first, "BHA Hrs Slid");
  const components = rows
    .map((row) => ({ serial: text(row, "Serial #"), itemCode: text(row, "Item Code"), description: text(row, "Description"), row }))
    .filter((row) => row.serial !== "" || row.itemCode !== "" || row.description !== "")
    .map((row, index) => ({
      index: index + 1,
      serial: row.serial,
      itemCode: row.itemCode,
      description: row.description,
      hrsSlid,
      hrsRot: num(first, "BHA Hrs Rot"),
      hrsTotal,
      bha,
    }));
  return {
    bha,
    section: text(first, "Section"),
    status: text(first, "Status"),
    motor: text(first, "Motor"),
    guidance: text(first, "Guidance"),
    metres,
    hrsTotal,
    hrsSlid,
    hrsRot: num(first, "BHA Hrs Rot"),
    hrsCirc: num(first, "BHA Hrs Circ"),
    hrsBelowRot: num(first, "BHA Below Rot"),
    activatedOn: text(first, "Activated On"),
    completedOn: text(first, "Completed On"),
    slideFraction: hrsTotal > 0 ? hrsSlid / hrsTotal : 0,
    components,
  };
}

interface InventoryItem {
  category: string;
  subCategory: string;
  hrsSinceService: number;
}

/** _OC_Inventory keyed by the upper-cased serial, the lookup CollectBhaTabSerialRows falls back to. */
function inventoryBySerial(book: SheetBook): Map<string, InventoryItem> {
  const out = new Map<string, InventoryItem>();
  for (const row of rowsOf(book, INV_SHEET)) {
    const serial = text(row, "SerialNumber").toUpperCase();
    if (serial === "") continue;
    out.set(serial, { category: text(row, "Category"), subCategory: text(row, "SubCategory"), hrsSinceService: num(row, "HrsSinceService") });
  }
  return out;
}

/**
 * CollectBhaTabSerialRows: one row per serial per BHA. DD tools stay, MWD tools go, and a tool
 * with no category stays only when its sub-description starts with "DD". Hours since service
 * come from the BHA row, then from the inventory. Sorted by BHA descending, then HSLS descending.
 */
export function serialHours(book: SheetBook): SerialHoursRow[] {
  const all = rowsOf(book, BHA_SHEET);
  // The BHA's own circ hours live on its first row; a component row leaves the cell blank.
  const circByBha = new Map<number, number>();
  for (const row of all) {
    const bha = bhaNumber(row);
    const circ = num(row, "BHA Hrs Circ");
    if (bha != null && circ !== 0 && !circByBha.has(bha)) circByBha.set(bha, circ);
  }
  const inventory = inventoryBySerial(book);
  const seen = new Set<string>();
  const out: SerialHoursRow[] = [];
  for (const row of all) {
    const bha = bhaNumber(row);
    const serial = text(row, "Serial #");
    if (bha == null || serial === "") continue;
    const key = serial.toUpperCase();
    const item = inventory.get(key);
    const category = (text(row, "Category") || item?.category || "").toUpperCase();
    const sub = (text(row, "Sub Description") || item?.subCategory || "").toUpperCase();
    const keep = category === "DD" || (category !== "MWD" && sub.startsWith("DD"));
    if (!keep || seen.has(`${key}|${bha}`)) continue;
    seen.add(`${key}|${bha}`);
    const hsls = num(row, "HSLS") || item?.hrsSinceService || 0;
    out.push({ serial, description: text(row, "Description"), circHours: num(row, "BHA Hrs Circ") || circByBha.get(bha) || 0, hsls, bha });
  }
  out.sort((a, b) => b.bha - a.bha || b.hsls - a.hsls);
  return out;
}

const HEADER_FILL = "#D9D9D9";
const BAND_FILL = "#E2EFDA";
const GRID = "#000000";
/** Excel character-widths from BuildDataSerialHoursTable, in pixels, so a description does not spill. */
const COLUMN_PX: Array<[number, number]> = [[24, 150], [25, 360], [26, 92], [27, 78], [28, 58]];

function paintRow(book: SheetBook, row: number, fill: string | null, descriptionLeft: boolean): void {
  const area = { row, col: 24, row2: row, col2: 28 };
  book.setLook(DATA, area, { bold: true, hAlign: "center", vAlign: "middle", fontFamily: "Calibri", fontSize: 11, wrap: "clip" });
  if (descriptionLeft) book.setLook(DATA, { row, col: 25, row2: row, col2: 25 }, { bold: true, hAlign: "left", vAlign: "middle", fontFamily: "Calibri", fontSize: 11, wrap: "clip" });
  book.setFill(DATA, area, fill);
  book.setFontColor(DATA, area, "#000000");
  for (let col = 24; col <= 28; col += 1) book.setBorder(DATA, { row, col, row2: row, col2: col }, "outside", "thin", GRID);
}

/**
 * BuildDataSerialHoursTable: Data!X2:AB, header on row 2, one row per serial from row 3.
 * Clears the whole block first so a shorter table never leaves the previous job's rows behind.
 * Each BHA is one band, hours are 0.00, and the description sits left, the same as the sheet in Excel.
 */
export function writeSerialHoursTable(book: SheetBook, rows: SerialHoursRow[]): number {
  if (!book.hasSheet(DATA)) return 0;
  book.clear(DATA, { row: SERIAL_HEADER_ROW, col: 24, row2: SERIAL_CLEAR_TO, col2: 28 }, true);
  for (const [col, px] of COLUMN_PX) book.setColumnWidth(DATA, col, px);
  const headers = ["Serial", "Description", "Circ Hours", "HSLS", "BHA#"];
  headers.forEach((header, index) => book.setValue(DATA, SERIAL_HEADER_ROW, 24 + index, header));
  paintRow(book, SERIAL_HEADER_ROW, HEADER_FILL, false);
  let band = false;
  let previous = -1;
  rows.forEach((row, index) => {
    const at = SERIAL_FIRST_ROW + index;
    if (row.bha !== previous) {
      if (previous >= 0) band = !band;
      previous = row.bha;
    }
    const values: CellScalar[] = [row.serial, row.description, row.circHours, row.hsls, row.bha];
    values.forEach((value, col) => book.setValue(DATA, at, 24 + col, value));
    book.setFormat(DATA, at, 26, "0.00");
    book.setFormat(DATA, at, 27, "0.00");
    book.setFormat(DATA, at, 28, "0");
    paintRow(book, at, band ? BAND_FILL : null, true);
  });
  return rows.length;
}
