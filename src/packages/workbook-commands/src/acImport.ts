import { checkAcReport } from "./acCheck";
import { AC_DETAIL_SHEET, parseAcDetail, type AcDetailStation } from "./acDetail";
import { unreadablePdfText } from "./glyphText";
import { cleanNum, SETUP_SHEET, showImportPath, SURVEY_SHEET } from "./planImport";
import { cellText, numeric, rgb, type CellArea, type SheetBook } from "./sheetBook";

// Port of MDL_Setup Import AC: the Compass Anticollision Report Summary rows to _OC_AC,
// the Setup J11:R16 table (BuildAcTable) and Data "AC Info & Concerns" (WriteAcConcerns).

export const AC_SHEET = "_OC_AC";
const AC_CHECK_COL = 6;
const DATA_SHEET = "Data";
export const AC_SF_RED = 1.5;
export const AC_SF_YELLOW = 2;
const MAX_HITS = 200;
const AC_BASE_ROW = 11;
const AC_BASE_COL = 10;
const AC_COLS = 9;
const AC_MAX = 5;

const WHITE = rgb(255, 255, 255);
const BAND_GREY = rgb(242, 242, 242);
const MID_GREY = rgb(217, 217, 217);
const DARK = rgb(89, 89, 89);
const BLACK = rgb(0, 0, 0);
const RED_TEXT = rgb(160, 30, 30);
const AMBER_TEXT = rgb(140, 90, 0);
const ROW_RULE = rgb(235, 235, 235);
const AC_RED = rgb(255, 199, 206);
const AC_YELLOW = rgb(255, 235, 156);
const AC_GREEN = rgb(198, 239, 206);

export interface AcHit {
  well: string;
  refMd: number;
  between: number;
  sf: number;
}

function findText(text: string, find: string, from = 0): number {
  return text.toLowerCase().indexOf(find.toLowerCase(), from);
}

function vbTrim(text: string): string {
  return text.replace(/^ +| +$/g, "");
}

export function cleanWellName(name: string): string {
  let out = name;
  while (out.includes("  ")) out = out.replaceAll("  ", " ");
  return vbTrim(out);
}

/** AcStoreHit: one slot per well; a later hit for the same name replaces the earlier one. */
function storeHit(hits: AcHit[], well: string, refMd: number, between: number, sf: number): void {
  if (sf <= 0) return;
  const name = cleanWellName(well);
  if (name === "") return;
  const same = hits.find((hit) => hit.well.toLowerCase() === name.toLowerCase());
  if (same) {
    same.refMd = refMd;
    same.between = between;
    same.sf = sf;
    return;
  }
  if (hits.length > MAX_HITS) return;
  hits.push({ well: name, refMd, between, sf });
}

/** ParseAcSummary: Summary rows with 0 < SF <= maxSF, one per offset well. */
export function parseAcSummary(text: string, maxSF = 2): AcHit[] {
  const hits: AcHit[] = [];
  let start = findText(text, "Summary");
  if (start < 0) start = 0;
  let end = text.length;
  for (const marker of ["Semi Major", "Offset Design", "Highside"]) {
    const p = findText(text, marker, start);
    if (p >= 0 && p < end) end = p;
  }
  const region = text.slice(start, end);

  // The extractor may put stray spaces inside numbers ("1, 110.00", "4. 11").
  const num2 = String.raw`\d{1,3}(?:,\s?\d{3})*\.\s?\d{2}`;
  const sfNum = String.raw`\d{1,3}\.\s?\d{3}`;
  const re = new RegExp(String.raw`(\d{2,3}/\S{5,30}/\d{1,2}\s*-\s*.{1,60}?)\s+(${num2})\s+(${num2})\s+(${num2})\s+(${num2})\s+(${sfNum})`, "g");
  const matches = [...region.matchAll(re)];
  for (const m of matches) {
    let sf = cleanNum(m[6]!);
    // Compass prints CC, ES and SF as separate depths for one offset. Only the SF line alerts.
    let tag = region.substr(m.index + m[0].length, 24);
    const lf = tag.indexOf("\n");
    if (lf >= 0) tag = tag.slice(0, lf);
    const cr = tag.indexOf("\r");
    if (cr >= 0) tag = tag.slice(0, cr);
    tag = vbTrim(tag).toUpperCase();
    if ((tag.startsWith("CC") || tag.startsWith("ES")) && !tag.includes("SF")) sf = 0;
    if (sf > 0 && sf <= maxSF) storeHit(hits, m[1]!, cleanNum(m[2]!), cleanNum(m[4]!), sf);
  }
  if (hits.length > 0 || matches.length > 0) return hits;

  // Legacy: "Level 4 , SF 2,604.552,604.1018.567.201.634 OFFSET WELL..." with run-together digits.
  let inSummary = false;
  for (const raw of text.split("\n")) {
    const line = vbTrim(raw);
    if (findText(line, "Summary") >= 0) inSummary = true;
    if (!inSummary) continue;
    if (findText(line, "Level") < 0 || findText(line, ", SF") < 0) continue;
    const mds = [...line.matchAll(/\d{1,3}(?:,\d{3})+\.\d{2}/g)];
    if (mds.length < 2) continue;
    const refMd = Number(mds[0]![0].replaceAll(",", ""));
    const offsetMd = mds[1]!;
    const after = line.slice(offsetMd.index + offsetMd[0].length);
    const trip = /(\d{1,2})\.(\d{2})(\d{1,2})\.(\d{2})(\d)\.(\d{3})/.exec(after);
    if (!trip) continue;
    const between = Number(`${trip[1]}.${trip[2]}`);
    const sf = Number(`${trip[5]}.${trip[6]}`);
    if (sf > 0 && sf <= maxSF) {
      let tail = vbTrim(after.slice(trip.index + trip[0].length));
      if (tail.length > 60) tail = tail.slice(0, 60);
      storeHit(hits, tail, refMd, between, sf);
      if (hits.length > MAX_HITS) break;
    }
  }
  return hits;
}

/** AcKeepsPlace: severity = lowest SF first (closer centres break ties); depth = shallowest first. */
function keepsPlace(a: AcHit, b: AcHit, byDepth: boolean): boolean {
  if (byDepth) {
    if (a.refMd < b.refMd) return true;
    if (a.refMd > b.refMd) return false;
    return a.sf <= b.sf;
  }
  if (a.sf < b.sf) return true;
  if (a.sf > b.sf) return false;
  return a.between <= b.between;
}

/** SortAcRows: the VBA's stable insertion sort. */
export function sortAcRows(rows: AcHit[], byDepth: boolean): AcHit[] {
  const out = [...rows];
  for (let i = 1; i < out.length; i += 1) {
    const key = out[i]!;
    let j = i - 1;
    while (j >= 0 && !keepsPlace(out[j]!, key, byDepth)) {
      out[j + 1] = out[j]!;
      j -= 1;
    }
    out[j + 1] = key;
  }
  return out;
}

// ---------------------------------------------------------------- Setup J11:R16

/** AcSetupCell: logical column 0..4 -> J, K:L, M:N, O:P, Q:R. */
function acSetupCell(r: number, logical: number): CellArea {
  const spans: Array<[number, number]> = [[10, 10], [11, 12], [13, 14], [15, 16], [17, 18]];
  const [col, col2] = spans[logical] ?? [10, 10];
  return { row: r, col, row2: r, col2 };
}

function mergeRow(book: SheetBook, r: number): void {
  for (let c = 0; c <= 4; c += 1) book.merge(SETUP_SHEET, acSetupCell(r, c));
}

function riskOf(sf: number): { risk: string; color: string } {
  if (sf < 1) return { risk: "CRITICAL", color: RED_TEXT };
  if (sf < 1.5) return { risk: "HIGH", color: RED_TEXT };
  return { risk: "CAUTION", color: AMBER_TEXT };
}

/** BuildAcTable: the severity-sorted hits, five at most, on Setup J11:R16. */
export function buildAcTable(book: SheetBook, hits: AcHit[]): void {
  if (!book.hasSheet(SETUP_SHEET)) return;
  const lastCol = AC_BASE_COL + AC_COLS - 1;
  const block = { row: AC_BASE_ROW, col: AC_BASE_COL, row2: AC_BASE_ROW + AC_MAX, col2: lastCol };
  book.unmerge(SETUP_SHEET, block);
  book.clear(SETUP_SHEET, block, true);
  const legacy = { row: 1, col: 19, row2: 251, col2: 23 };
  book.unmerge(SETUP_SHEET, legacy);
  book.clear(SETUP_SHEET, legacy, true);

  let r = AC_BASE_ROW;
  mergeRow(book, r);
  ["#", "Ref MD (m)", "Between Centres (m)", "Separation Factor", "Risk"].forEach((head, c) => {
    const cell = acSetupCell(r, c);
    book.setValue(SETUP_SHEET, cell.row, cell.col, head);
    book.setFill(SETUP_SHEET, cell, MID_GREY);
    book.setFontColor(SETUP_SHEET, cell, BLACK);
    book.setLook(SETUP_SHEET, cell, { bold: true, fontFamily: "Consolas", fontSize: 8, vAlign: "middle", hAlign: c === 0 || c === 4 ? "center" : "right" });
  });
  book.setBorder(SETUP_SHEET, { row: r, col: AC_BASE_COL, row2: r, col2: lastCol }, "bottom", "thin", MID_GREY);
  r += 1;

  const shown = hits.slice(0, AC_MAX);
  if (shown.length === 0) {
    const row = { row: r, col: AC_BASE_COL, row2: r, col2: lastCol };
    book.merge(SETUP_SHEET, row);
    book.setValue(SETUP_SHEET, r, AC_BASE_COL, "  No data  --  use Import AC button to load");
    book.setFill(SETUP_SHEET, row, BAND_GREY);
    book.setFontColor(SETUP_SHEET, row, DARK);
    book.setLook(SETUP_SHEET, row, { fontFamily: "Consolas", fontSize: 8, vAlign: "middle" });
    r += 1;
  } else {
    shown.forEach((hit, i) => {
      const band = i % 2 === 0 ? WHITE : BAND_GREY;
      const { risk, color } = riskOf(hit.sf);
      const values: Array<string | number> = [i + 1, hit.refMd, hit.between, hit.sf, risk];
      mergeRow(book, r);
      values.forEach((value, c) => {
        const cell = acSetupCell(r, c);
        book.setValue(SETUP_SHEET, cell.row, cell.col, value);
        book.setFill(SETUP_SHEET, cell, band);
        book.setFontColor(SETUP_SHEET, cell, c === 3 || c === 4 ? color : BLACK);
        book.setLook(SETUP_SHEET, cell, { bold: c === 4, fontFamily: "Consolas", fontSize: 8, vAlign: "middle", hAlign: c === 0 || c === 4 ? "center" : "right" });
        if (c === 3) book.setFormat(SETUP_SHEET, cell.row, cell.col, "0.000");
        if (c === 1 || c === 2) book.setFormat(SETUP_SHEET, cell.row, cell.col, "0.00");
      });
      book.setBorder(SETUP_SHEET, { row: r, col: AC_BASE_COL, row2: r, col2: lastCol }, "bottom", "hair", ROW_RULE);
      r += 1;
    });
  }
  const endRow = shown.length === 0 ? AC_BASE_ROW + AC_MAX : r - 1;
  book.setBorder(SETUP_SHEET, { row: AC_BASE_ROW, col: AC_BASE_COL, row2: endRow, col2: lastCol }, "outside", "thin", MID_GREY);
}

// ---------------------------------------------------------------- Data "AC Info & Concerns"

function bandColor(sf: number): string {
  if (sf < AC_SF_RED) return AC_RED;
  if (sf < AC_SF_YELLOW) return AC_YELLOW;
  return AC_GREEN;
}

function findInColumnB(book: SheetBook, text: string, fromRow: number, toRow: number): number {
  for (let r = fromRow; r <= toRow; r += 1) if (findText(cellText(book.value(DATA_SHEET, r, 2)), text) >= 0) return r;
  return 0;
}

/** The Data tab rows "AC Info & Concerns" is allowed to fill. 0 when that table is not there. */
export function acConcernCapacity(book: SheetBook): number {
  const span = acConcernSpan(book);
  return span ? span.last - span.first + 1 : 0;
}

/**
 * The wells the Data tab shows: the lowest separation factors, as many as the table holds,
 * then listed by depth. `hits` must already be severity-sorted, the way Import AC stores them.
 */
export function concernWells(hits: AcHit[], capacity: number): AcHit[] {
  const room = Math.max(0, Math.min(hits.length, capacity));
  return sortAcRows(hits.slice(0, room), true);
}

function acConcernSpan(book: SheetBook): { first: number; last: number } | null {
  if (!book.hasSheet(DATA_SHEET)) return null;
  const title = findInColumnB(book, "AC Info & Concerns", 1, 200);
  if (title === 0) return null;
  const first = title + 2;
  let last = first + 7;
  const bottom = findInColumnB(book, "Motors On Location", first, first + 40);
  if (bottom > first) last = bottom - 1;
  if (last < first) return null;
  return { first, last };
}

function readPlanPoints(book: SheetBook): Array<{ md: number; n: number; e: number; tvd: number }> {
  if (!book.hasSheet(SURVEY_SHEET)) return [];
  const last = book.lastUsedRow(SURVEY_SHEET, 1);
  const out: Array<{ md: number; n: number; e: number; tvd: number }> = [];
  for (let row = 3; row <= last; row += 1) {
    const md = numeric(book.value(SURVEY_SHEET, row, 1));
    const tvd = numeric(book.value(SURVEY_SHEET, row, 4));
    const n = numeric(book.value(SURVEY_SHEET, row, 5));
    const e = numeric(book.value(SURVEY_SHEET, row, 6));
    if (md == null || tvd == null || n == null || e == null) continue;
    out.push({ md, n, e, tvd });
  }
  return out;
}

/** The Summary rows Import AC wrote, in the sheet's severity order. */
export function readAcHits(book: SheetBook): AcHit[] {
  if (!book.hasSheet(AC_SHEET)) return [];
  const last = book.lastUsedRow(AC_SHEET, 1);
  const hits: AcHit[] = [];
  for (let row = 3; row <= last; row += 1) {
    const well = cellText(book.value(AC_SHEET, row, 1)).trim();
    const refMd = numeric(book.value(AC_SHEET, row, 2));
    const between = numeric(book.value(AC_SHEET, row, 3));
    const sf = numeric(book.value(AC_SHEET, row, 4));
    if (well === "" || refMd == null || between == null || sf == null || sf <= 0) continue;
    hits.push({ well, refMd, between, sf });
  }
  return hits;
}

function writeAcDetail(book: SheetBook, stations: AcDetailStation[]): void {
  book.resetSheet(AC_DETAIL_SHEET);
  ["Well", "Ref MD", "Off MD", "Off TVD", "N", "E", "C2C", "SF", "Warning"].forEach((head, i) => {
    book.setValue(AC_DETAIL_SHEET, 2, i + 1, head);
  });
  stations.forEach((station, i) => {
    const row = i + 3;
    const values: Array<string | number | null> = [
      station.well, station.refMd, station.offMd, station.offTvd, station.n, station.e, station.c2c, station.sf,
      station.warn === "" ? null : station.warn,
    ];
    values.forEach((value, col) => book.setValue(AC_DETAIL_SHEET, row, col + 1, value));
  });
}

/** WriteAcConcerns. Returns the MsgBox text when the table cannot hold every well. */
export function writeAcConcerns(book: SheetBook, hits: AcHit[], quiet = false): string | null {
  const span = acConcernSpan(book);
  if (!span) return null;
  const firstRow = span.first;
  const lastRow = span.last;

  if (lastRow >= firstRow) {
    book.clear(DATA_SHEET, { row: firstRow, col: 2, row2: lastRow, col2: 6 });
    for (let r = firstRow; r <= lastRow; r += 1) book.setFill(DATA_SHEET, { row: r, col: 2, row2: r, col2: 6 }, (r - firstRow) % 2 === 0 ? null : BAND_GREY);
  }
  const capacity = lastRow - firstRow + 1;
  // Severity picks who gets a slot; the slots then read in drilling order.
  const shown = concernWells(hits, capacity);
  shown.forEach((hit, i) => {
    const r = firstRow + i;
    const name = book.mergeAt(DATA_SHEET, r, 2);
    book.setValue(DATA_SHEET, name.row, name.col, hit.well);
    book.setValue(DATA_SHEET, r, 4, hit.sf);
    book.setFormat(DATA_SHEET, r, 4, "0.000");
    book.setValue(DATA_SHEET, r, 5, hit.between);
    book.setFormat(DATA_SHEET, r, 5, "0.00");
    book.setValue(DATA_SHEET, r, 6, hit.refMd);
    book.setFormat(DATA_SHEET, r, 6, "0.00");
    book.setFill(DATA_SHEET, { row: r, col: 2, row2: r, col2: 6 }, bandColor(hit.sf));
  });
  if (hits.length > capacity && !quiet) {
    return `AC import read ${hits.length} offset well(s); this table holds ${capacity}.\n\nKept the ${capacity} with the lowest separation factors and listed them by depth.\nThe full list is on the Setup sheet AC table.`;
  }
  return null;
}

// ---------------------------------------------------------------- Import AC

export interface AcImportResult {
  ok: boolean;
  alert?: string;
  /** "AC Info & Concerns" overflow note, shown after the import. */
  note?: string;
  status: string;
  hits: AcHit[];
  /** The frame-check line stored on `_OC_AC` and shown in the 3D view. */
  badge?: string;
}

export const NO_AC_TEXT = "Could not read any text from that PDF. A scanned (image-only) report has no text to import.\n\nPath stored. Try again with the Compass Anticollision Report PDF.";
export const AC_NO_LETTER_MAP = "That PDF's text cannot be read: its fonts store glyph numbers instead of letters, and the glyphs did not match Arial. Nothing was imported.\n\nPath stored. Ask for the anticollision report printed to the Adobe PDF printer.";
export const AC_NO_WELLS = "The AC report was read, but its Summary table gave no offset wells, so the AC table is empty.\n\nCheck that this PDF is the Compass Anticollision Report.";

/** ImportAntiCollisionFile. maxSF defaults high so every offset well is read, then sorted and banded. */
export function importAntiCollision(book: SheetBook, path: string, text: string, acRowAnchor: { row: number; col: number } | null, maxSF = 1_000_000, quiet = false): AcImportResult {
  book.resetSheet(AC_SHEET);
  book.setValue(AC_SHEET, 1, 1, path);
  if (text.trim() === "") {
    writeAcDetail(book, []);
    showImportPath(book, acRowAnchor, AC_SHEET);
    return { ok: false, alert: NO_AC_TEXT, status: "No text in that PDF. Nothing was imported.", hits: [] };
  }
  const hits = sortAcRows(parseAcSummary(text, maxSF), false);
  if (hits.length === 0 && unreadablePdfText(text)) {
    writeAcDetail(book, []);
    showImportPath(book, acRowAnchor, AC_SHEET);
    return { ok: false, alert: AC_NO_LETTER_MAP, status: "No readable text in that PDF. Nothing was imported.", hits: [] };
  }
  ["Offset Well - Wellbore - Design", "Ref MD (m)", "Between Centres (m)", "Separation Factor"].forEach((head, i) => book.setValue(AC_SHEET, 2, i + 1, head));
  hits.forEach((hit, i) => {
    book.setValue(AC_SHEET, i + 3, 1, hit.well);
    book.setValue(AC_SHEET, i + 3, 2, hit.refMd);
    book.setValue(AC_SHEET, i + 3, 3, hit.between);
    book.setValue(AC_SHEET, i + 3, 4, hit.sf);
  });
  const stations = parseAcDetail(text);
  writeAcDetail(book, stations);
  buildAcTable(book, hits);
  const overflow = writeAcConcerns(book, hits, quiet);
  const shown = concernWells(hits, acConcernCapacity(book));
  const check = checkAcReport(shown, stations, readPlanPoints(book));
  book.setValue(AC_SHEET, 1, AC_CHECK_COL, check.badge);
  const note = hits.length === 0 && !quiet ? AC_NO_WELLS : overflow;
  showImportPath(book, acRowAnchor, AC_SHEET);
  const concerns = hits.filter((hit) => hit.sf < AC_SF_YELLOW).length;
  return {
    ok: true,
    ...(note ? { note } : {}),
    status: `AC import complete: ${hits.length} offset well(s), ${concerns} with SF < ${AC_SF_YELLOW.toFixed(1)}.`,
    hits,
    badge: check.badge,
  };
}
