import { cellText, fixed2, isErrorValue, isNumber, numeric, SheetBook } from "./sheetBook";
import { SS_SHEET } from "./slidesheet";

// Port of MDL_PipeTally + frmPipeTally: pipe lengths from the rig CSV into F7:F315 of the
// Pipe Tally Calculator, then the calculated bit depths H7:H315 appended to Slidesheet D.

export const PT_SHEET = "Pipe Tally Calculator";
const PT_FIRST = 7;
const PT_LAST = 315;
const PT_CAP = PT_LAST - PT_FIRST + 1;
const MIN_JOINT_M = 0.1;
const MAX_JOINT_M = 40;

export const TALLY_LAYOUT_HELP = [
  "The calculator reads and writes these cells on that sheet:",
  "    B5        start depth (last bit depth on the Slidesheet)",
  "    C11       stands per joint",
  "    D11       split length",
  `    F${PT_FIRST}:F${PT_LAST}   pipe joint lengths (input)`,
  `    H${PT_FIRST}:H${PT_LAST}   calculated bit depths (output)`,
  "",
  "Restore that sheet from a copy in the backups folder, then try again.",
].join("\n");

export function missingTallySheet(book: SheetBook): string | null {
  return book.hasSheet(PT_SHEET) ? null : `Worksheet "${PT_SHEET}" does not exist in this workbook.`;
}

/** Last positive Slidesheet D in D12:D505. */
export function lastSlidesheetBitMd(book: SheetBook): number | null {
  for (let r = 505; r >= 12; r -= 1) {
    const v = book.value(SS_SHEET, r, 4);
    if (isErrorValue(v) || v == null || v === "") continue;
    const n = numeric(v);
    if (n != null && n > 0) return n;
  }
  return null;
}

function csvField(line: string, index1: number): string {
  if (line === "") return "";
  const parts = line.split(",");
  if (index1 < 1 || index1 - 1 > parts.length - 1) return "";
  let s = parts[index1 - 1]!.trim();
  if (s.startsWith("\"")) s = s.slice(1);
  if (s.endsWith("\"") && s.length > 0) s = s.slice(0, -1);
  return s;
}

function vbaVal(text: string): number {
  const match = /^\s*[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?/i.exec(text);
  return match ? Number(match[0]) : 0;
}

export type CsvLengths = { ok: true; lengths: number[]; note: string } | { ok: false; error: string };

/** LoadPipeLengthsFromCsv: column G total matched to the bit, then column F lengths below it. */
export function lengthsFromCsv(text: string, bitMd: number | null): CsvLengths {
  if (bitMd == null) return { ok: false, error: "No bit depth in Slidesheet column D." };
  if (text.length === 0) return { ok: false, error: "The file is empty." };
  const rows = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
  let bestRow = -1;
  let bestG = 0;
  let matched = false;
  for (let i = 0; i < rows.length; i += 1) {
    const gText = csvField(rows[i]!, 7);
    if (gText === "" || numeric(gText) == null) continue;
    const g = vbaVal(gText);
    if (Math.abs(g - bitMd) <= 0.05) {
      bestRow = i;
      bestG = g;
      matched = true;
      break;
    }
    if (g <= bitMd + 0.001 && (bestRow < 0 || g > bestG)) {
      bestRow = i;
      bestG = g;
    }
  }
  if (bestRow < 0) return { ok: false, error: `No tally total in column G is at or below the bit depth ${fixed2(bitMd)} m.` };
  if (!matched && bitMd - bestG > 20) {
    return { ok: false, error: `This file only reaches ${fixed2(bestG)} m in column G, and the bit is at ${fixed2(bitMd)} m.\nPick a newer pipe tally export.` };
  }
  const lengths: number[] = [];
  let capped = false;
  for (let i = bestRow + 1; i < rows.length; i += 1) {
    const fText = csvField(rows[i]!, 6);
    if (fText === "" || numeric(fText) == null) continue;
    const f = vbaVal(fText);
    if (f < MIN_JOINT_M || f > MAX_JOINT_M) continue;
    if (lengths.length >= PT_CAP) {
      capped = true;
      break;
    }
    lengths.push(f);
  }
  if (lengths.length === 0) {
    return { ok: false, error: `Bit depth ${fixed2(bitMd)} m matched tally total ${fixed2(bestG)} m. No single pipe lengths in column F below that row.` };
  }
  let note = matched ? "" : `Closest tally total at or below the bit is ${fixed2(bestG)} m (bit ${fixed2(bitMd)} m).`;
  if (capped) note = `${note} Stopped at ${lengths.length} joints (the form holds that many).`;
  return { ok: true, lengths, note };
}

/** The note ImportPipeTallyCsv leaves for the form. */
export function importNote(result: { lengths: number[]; note: string }): string {
  return result.note.length > 0
    ? `Imported ${result.lengths.length} pipe length(s). ${result.note}`
    : `Imported ${result.lengths.length} pipe length(s). Set stands and split, then add.`;
}

export function writeLengths(book: SheetBook, lengths: number[]): void {
  book.clear(PT_SHEET, `F${PT_FIRST}:F${PT_LAST}`);
  lengths.slice(0, PT_CAP).forEach((length, i) => book.setValue(PT_SHEET, PT_FIRST + i, 6, length));
}

export function clearTallyInputs(book: SheetBook): void {
  book.clear(PT_SHEET, `F${PT_FIRST}:F${PT_LAST}`);
}

export interface TallyView {
  start: string;
  stands: string;
  split: string;
  lengths: number[];
  depths: number[];
}

function nzText(value: unknown): string {
  return value == null ? "" : String(value);
}

/** ReloadFromSheet: start depth, settings, and the lengths / depths preview. */
export function readTally(book: SheetBook): TallyView {
  const start = numeric(book.value(PT_SHEET, 5, 2));
  const lengths: number[] = [];
  const depths: number[] = [];
  for (let r = PT_FIRST; r <= PT_LAST; r += 1) {
    const f = book.value(PT_SHEET, r, 6);
    const n = f == null || f === "" || isErrorValue(f) ? null : numeric(f);
    if (n != null && n > 0) lengths.push(n);
    const h = book.value(PT_SHEET, r, 8);
    if (isNumber(h)) depths.push(h);
  }
  return {
    start: start == null ? "-" : fixed2(start),
    stands: nzText(book.value(PT_SHEET, 11, 3)),
    split: nzText(book.value(PT_SHEET, 11, 4)),
    lengths,
    depths,
  };
}

/** SyncSettingsToSheet: blank or bad stands become 1; a blank split clears D11. */
export function writeSettings(book: SheetBook, standsText: string, splitText: string): { stands: string } {
  let stands = 1;
  if (standsText.trim() !== "" && numeric(standsText) != null) stands = Math.max(1, Math.round(vbaVal(standsText)));
  book.setValue(PT_SHEET, 11, 3, stands);
  if (splitText.trim() === "" || numeric(splitText) == null) book.clear(PT_SHEET, "D11");
  else book.setValue(PT_SHEET, 11, 4, vbaVal(splitText));
  return { stands: String(stands) };
}

export type AppendResult = { ok: true; message: string; warning: string | null; added: { row: number; row2: number } | null } | { ok: false; message: string };

/** AppendTallyBitDepths on a recalculated sheet. */
export function appendTallyBitDepths(book: SheetBook): AppendResult {
  const missing = missingTallySheet(book);
  if (missing) return { ok: false, message: `Cannot add bit depths.\n\n${missing}\n\n${TALLY_LAYOUT_HELP}` };
  let lastRow = 11;
  for (let r = 12; r <= 505; r += 1) if (cellText(book.value(SS_SHEET, r, 4)) !== "") lastRow = r;
  if (lastRow < 12) return { ok: false, message: "No bit depth entries found on Slidesheet column D (from row 12)." };
  const depths: number[] = [];
  for (let r = PT_FIRST; r <= PT_LAST; r += 1) {
    const v = book.value(PT_SHEET, r, 8);
    if (isNumber(v)) depths.push(v);
  }
  if (depths.length === 0) return { ok: false, message: `No tally bit depths found in H${PT_FIRST}:H${PT_LAST} to append.` };
  const destRow = lastRow + 1;
  let count = depths.length;
  let truncated = false;
  if (destRow + count - 1 > 505) {
    count = Math.max(0, 505 - destRow + 1);
    truncated = true;
  }
  for (let i = 0; i < count; i += 1) book.setValue(SS_SHEET, destRow + i, 4, depths[i]!);
  book.clear(PT_SHEET, `F${PT_FIRST}:F${PT_LAST}`);
  // The caller recalculates, then runs RefreshRotateMetres / RefreshSlideComments on `added`.
  const added = count > 0 ? { row: destRow, row2: destRow + count - 1 } : null;
  return {
    ok: true,
    message: `Added ${count} bit depth(s) to Slidesheet starting at D${destRow}. Pipe tally F7:F315 cleared.`,
    warning: truncated ? `Stopped after ${count} values: Slidesheet bit-depth rows only go through D505.` : null,
    added,
  };
}
