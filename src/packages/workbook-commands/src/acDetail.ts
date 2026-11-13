import { cleanNum } from "./planImport";
import { cellText, numeric, type SheetBook } from "./sheetBook";

export const AC_DETAIL_SHEET = "_OC_AC_Detail";

// Port of MDL_AcPoc_AC.AcPoc_ParseDetail and the well-name helpers it shares
// with MDL_AcPoc_Pdf. The text is the same pdftotext -layout text Import AC already uses.

export interface AcDetailStation {
  well: string;
  refMd: number;
  refTvd: number;
  offMd: number;
  offTvd: number;
  n: number;
  e: number;
  c2c: number;
  sf: number;
  warn: string;
}

const DETAIL_NUMBER = /-?\d{1,6}(?:,\s?\d{3})*\.\s?\d{2,3}/g;
const MAX_STATIONS = 20_000;

function cleanWellName(name: string): string {
  let out = name;
  while (out.includes("  ")) out = out.replaceAll("  ", " ");
  return out.replace(/^ +| +$/g, "");
}

/** AcPoc_CleanNum: the leading number, after commas and spaces are removed. */
export function vbaCleanNum(text: string): number {
  const compact = text.trim().replaceAll(",", "").replaceAll(" ", "");
  let out = "";
  for (const ch of compact) {
    if (/[0-9.\-]/.test(ch)) out += ch;
    else if (out.length > 0) break;
  }
  if (out === "" || out === "-" || out === "." || Number.isNaN(Number(out))) return 0;
  return Number(out);
}

/** AcPoc_WellKey: the wellbore id ending in the meridian, such as 100/01-02-003-04W5. */
export function wellKey(name: string): string {
  const text = cleanWellName(name);
  for (let i = 0; i <= text.length - 3; i += 1) {
    if (text[i]!.toUpperCase() !== "W" || text[i + 2] !== "/") continue;
    let start = i;
    while (start > 0 && /[0-9/\-]/.test(text[start - 1]!)) start -= 1;
    let end = i + 2;
    while (end < text.length && /[0-9A-Za-z]/.test(text[end]!)) end += 1;
    return text.slice(start, end).toUpperCase();
  }
  return text.toUpperCase();
}

/** AcPoc_WellsMatch. */
export function wellsMatch(a: string, b: string): boolean {
  const left = a.trim();
  const right = b.trim();
  if (left === "" || right === "") return false;
  if (right.toLowerCase().includes(left.toLowerCase())) return true;
  if (left.toLowerCase().includes(right.toLowerCase())) return true;
  const leftKey = wellKey(left);
  const rightKey = wellKey(right);
  return leftKey.length > 0 && leftKey === rightKey;
}

function extractNums(line: string): number[] {
  const out: number[] = [];
  for (const match of line.matchAll(DETAIL_NUMBER)) {
    out.push(cleanNum(match[0]));
    if (out.length > 50) break;
  }
  return out;
}

function columnOf(header: string, word: string): number {
  const at = header.toLowerCase().indexOf(word.toLowerCase());
  return at < 0 ? 0 : at + 1;
}

function sliceField(line: string, startCol: number, nextCol: number): string {
  if (startCol < 1) return "";
  const start = startCol - 1;
  return nextCol > startCol ? line.slice(start, nextCol - 1) : line.slice(start);
}

function warningOf(line: string): string {
  const tail = line.trim().slice(-2).toUpperCase();
  const upper = line.toUpperCase();
  let warn = "";
  if (upper.includes(" CC") || tail === "CC") warn = "CC";
  if (upper.includes(" ES") || tail === "ES") warn = "ES";
  if (upper.includes(" SF") || tail === "SF") warn = "SF";
  return warn;
}

function offsetDesignName(page: string): string {
  const at = page.toLowerCase().indexOf("offset design");
  if (at < 0) return "";
  let rest = page.slice(at + "offset design".length).replaceAll("\n", " ").replaceAll("\r", " ").trim();
  const cut = rest.toLowerCase().indexOf("offset site");
  if (cut >= 0) rest = rest.slice(0, cut);
  return cleanWellName(rest);
}

function stationFrom(line: string, nums: number[], columns: { n: number; e: number; bc: number; be: number; sf: number }): Pick<AcDetailStation, "n" | "e" | "c2c" | "sf"> | null {
  if (nums.length >= 13) return { n: nums[7]!, e: nums[8]!, c2c: nums[9]!, sf: nums[12]! };
  if (nums.length === 12) return { n: nums[6]!, e: nums[7]!, c2c: nums[8]!, sf: nums[11]! };
  if (columns.n > 0 && columns.e > 0 && columns.bc > 0) {
    return {
      n: vbaCleanNum(sliceField(line, columns.n, columns.e)),
      e: vbaCleanNum(sliceField(line, columns.e, columns.bc)),
      c2c: vbaCleanNum(sliceField(line, columns.bc, columns.be)),
      sf: vbaCleanNum(sliceField(line, columns.sf, 0)),
    };
  }
  if (nums.length < 10) return null;
  return { n: nums[7] ?? 0, e: nums[8] ?? 0, c2c: nums[9] ?? 0, sf: 0 };
}

/** AcPoc_ParseDetail. One station per Offset Design data row. */
export function parseAcDetail(text: string): AcDetailStation[] {
  const out: AcDetailStation[] = [];
  for (const page of text.split("\f")) {
    if (!page.toLowerCase().includes("offset design")) continue;
    const well = offsetDesignName(page);
    if (well === "") continue;
    const lines = page.split("\n");
    const headerAt = lines.findIndex((line) => line.includes("+N/-S") && line.toLowerCase().includes("centres"));
    const header = headerAt >= 0 ? lines[headerAt]! : "";
    const columns = {
      n: columnOf(header, "+N/-S"),
      e: columnOf(header, "+E/-W"),
      bc: columnOf(header, "Centres"),
      be: columnOf(header, "Ellipses"),
      sf: columnOf(header, "Factor"),
    };
    for (const line of lines) {
      if (line.includes("Offset Design") || line.toLowerCase().includes("page ")) continue;
      const nums = extractNums(line);
      if (nums.length < 10 || nums[0]! < 0 || nums[0]! > 12_000) continue;
      const picked = stationFrom(line, nums, columns);
      if (!picked) continue;
      out.push({
        well,
        refMd: nums[0]!,
        refTvd: nums[1]!,
        offMd: nums[2]!,
        offTvd: nums[3]!,
        n: picked.n,
        e: picked.e,
        c2c: picked.c2c,
        sf: picked.sf,
        warn: warningOf(line),
      });
      if (out.length >= MAX_STATIONS) return out;
    }
  }
  return out;
}

/** The detail stations Import AC wrote to `_OC_AC_Detail`. Reference TVD is not stored on the sheet. */
export function readAcDetail(book: SheetBook): AcDetailStation[] {
  if (!book.hasSheet(AC_DETAIL_SHEET)) return [];
  const last = book.lastUsedRow(AC_DETAIL_SHEET, 1);
  const out: AcDetailStation[] = [];
  for (let row = 3; row <= last; row += 1) {
    const well = cellText(book.value(AC_DETAIL_SHEET, row, 1)).trim();
    const refMd = numeric(book.value(AC_DETAIL_SHEET, row, 2));
    const offMd = numeric(book.value(AC_DETAIL_SHEET, row, 3));
    const offTvd = numeric(book.value(AC_DETAIL_SHEET, row, 4));
    const n = numeric(book.value(AC_DETAIL_SHEET, row, 5));
    const e = numeric(book.value(AC_DETAIL_SHEET, row, 6));
    const c2c = numeric(book.value(AC_DETAIL_SHEET, row, 7));
    const sf = numeric(book.value(AC_DETAIL_SHEET, row, 8));
    if (well === "" || refMd == null || offMd == null || offTvd == null || n == null || e == null || c2c == null || sf == null) continue;
    out.push({
      well,
      refMd,
      refTvd: 0,
      offMd,
      offTvd,
      n,
      e,
      c2c,
      sf,
      warn: cellText(book.value(AC_DETAIL_SHEET, row, 9)),
    });
  }
  return out;
}
