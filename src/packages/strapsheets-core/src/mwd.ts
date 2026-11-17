// The MWD Carrier: four subs strapped together, entered in rounds (every SN, then every length, then
// every OD). The sheet keeps each sub's own length and works out the carrier's total length and
// largest OD itself; this module does the same, so the phone's confirm screen and its card agree
// with the sheet's MWD.P.

import { decimalText, type EntryPayload } from "./entry";
import type { EntryFamily, RoundStep } from "./schema";
import type { Cells } from "./wizard";

export const MWD_WHICH = "MWD Carrier";

/**
 * Where each value sits in the card's row, as apps-script/mwd_carrier.js lays out MWD_COL:
 * 0 date, 1 carrier, then four columns per round, ID, the BOX/PIN and thread ends, total, max OD.
 */
const FIRST_ROUND_COL = 2;
const SUBS = 4;
const ID_COL = 14;
const END_COLS = { boxUpper: 15, boxLower: 16, threadUpper: 17, threadLower: 18 } as const;
const TOTAL_COL = 19;
const MAX_OD_COL = 20;

export interface MwdRounds {
  sn: RoundStep;
  length: RoundStep;
  od: RoundStep;
}

/** The three rounds of the MWD family, or null when the family is not one. */
export function mwdRounds(family: EntryFamily): MwdRounds | null {
  const rounds = family.steps.filter((step): step is RoundStep => step.kind === "round");
  const sn = rounds.find((round) => !round.decimal);
  const length = rounds.find((round) => round.cumulative);
  const od = rounds.find((round) => round.decimal && !round.cumulative);
  return sn && length && od ? { sn, length, od } : null;
}

/** A sheet number: what a number cell shows, so "152.40" prints as 152.4. */
function sheetNumber(text: string): string {
  const parsed = decimalText(text);
  return parsed == null || parsed === "" ? "" : String(Number(parsed));
}

/**
 * Running totals from the pin (A, A+B, ...) to each sub's own length. A blank is skipped and the next
 * total is measured from the last one entered, as the sheet does when it stores the row.
 */
export function mwdSubLengths(totals: string[]): string[] {
  let previous = 0;
  return totals.map((raw) => {
    const text = decimalText(raw);
    if (text == null || text === "") return "";
    const value = Number(text);
    const own = Math.round((value - previous) * 1000) / 1000;
    previous = value;
    return String(own);
  });
}

export interface MwdSummary {
  /** Each sub's own length, "" where the total was skipped. */
  subLengths: string[];
  /** The carrier's total length: the sum of its sub lengths. */
  total: string;
  /** The largest OD entered. */
  maxOd: string;
}

export function mwdSummary(family: EntryFamily, cells: Cells): MwdSummary {
  const rounds = mwdRounds(family);
  if (!rounds) return { subLengths: [], total: "", maxOd: "" };
  const subLengths = mwdSubLengths(rounds.length.cols.map((sub) => cells[sub.col] ?? ""));
  const lengths = subLengths.filter((value) => value !== "").map(Number);
  const ods = rounds.od.cols.map((sub) => sheetNumber(cells[sub.col] ?? "")).filter((value) => value !== "").map(Number);
  return {
    subLengths,
    total: lengths.length > 0 ? String(Math.round(lengths.reduce((sum, value) => sum + value, 0) * 1000) / 1000) : "",
    maxOd: ods.length > 0 ? String(Math.max(...ods)) : "",
  };
}

/** The date the card prints: the day the entry was taken, on this phone's clock. */
export function mwdDate(at: string | undefined): string {
  if (!at) return "";
  const when = new Date(at);
  if (Number.isNaN(when.getTime())) return "";
  const pad = (value: number) => String(value).padStart(2, "0");
  return `${when.getFullYear()}-${pad(when.getMonth() + 1)}-${pad(when.getDate())}`;
}

function stepCell(family: EntryFamily, id: string): string | number | null {
  const step = family.steps.find((candidate) => candidate.id === id);
  return step && (step.kind === "text" || step.kind === "decimal") ? step.col : null;
}

/**
 * The entry as the card's row: the named cells moved to the column each card cell reads, own lengths
 * in place of running totals, and the total and largest OD worked out.
 */
export function mwdCardEntry(family: EntryFamily, entry: EntryPayload): EntryPayload {
  const rounds = mwdRounds(family);
  if (!rounds) return entry;
  const { cells } = entry;
  const row: Record<string, string> = {};
  const put = (column: number, value: string) => {
    if (value !== "") row[String(column)] = value;
  };
  const serial = stepCell(family, "serial");
  const id = stepCell(family, "id");
  put(0, mwdDate(entry.at));
  put(1, serial == null ? "" : (cells[serial] ?? "").trim().toUpperCase());
  rounds.sn.cols.slice(0, SUBS).forEach((sub, index) => put(FIRST_ROUND_COL + index, (cells[sub.col] ?? "").trim().toUpperCase()));
  const summary = mwdSummary(family, cells);
  summary.subLengths.slice(0, SUBS).forEach((value, index) => put(FIRST_ROUND_COL + SUBS + index, value));
  rounds.od.cols.slice(0, SUBS).forEach((sub, index) => put(FIRST_ROUND_COL + 2 * SUBS + index, sheetNumber(cells[sub.col] ?? "")));
  put(ID_COL, id == null ? "" : sheetNumber(cells[id] ?? ""));
  for (const step of family.steps) {
    if (step.kind === "boxPin") {
      put(END_COLS.boxUpper, (cells[step.cols.upper] ?? "").trim());
      put(END_COLS.boxLower, (cells[step.cols.lower] ?? "").trim());
    }
    if (step.kind === "threads") {
      put(END_COLS.threadUpper, (cells[step.cols.upper] ?? "").trim());
      put(END_COLS.threadLower, (cells[step.cols.lower] ?? "").trim());
    }
  }
  put(TOTAL_COL, summary.total);
  put(MAX_OD_COL, summary.maxOd);
  return { which: entry.which, cells: row };
}
