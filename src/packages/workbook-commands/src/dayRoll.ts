import { slideMetersBetween } from "../../dd-math/src/project";
import { columnIndex, readColumn, readNumber, sheetOf, type SnapshotBook } from "./cells";
import { excelRound } from "./sheetBook";

/** Metres slid inside one Data day row, clipped to that row's start and end depth. */
export function daySlideMetres(book: SnapshotBook, row: number): number {
  const from = readNumber(book, "Data", `I${row}`);
  const to = readNumber(book, "Data", `J${row}`);
  if (from == null || to == null) return 0;
  return slideMetersBetween(
    from,
    to,
    readColumn(book, "Slidesheet", "R", 13, 505),
    readColumn(book, "Slidesheet", "S", 13, 505),
  );
}

// Port of MDL_DayRoll + frmDayHours: picking a period in Data!H4:H22 fills the row's
// start depth and slide metres/time, then asks for End Depth and the Pason running totals.

export const DAY_FIRST = 4;
export const DAY_LAST = 22;
const DAY_TOTALS = 24;
export const DAY_PERIOD_COLUMN = "H";
const PERIODS = ["6am", "4pm", "Midnight", "End of Run"] as const;
const SRC_FIRST = 36;
const SRC_LAST = 47;
const SS_FIRST = 13;
const SS_LAST = 505;
const ERROR_TEXT = /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|SPILL!|CALC!|GETTING_DATA)$/;

export type DayWrite = { a1: string; value: number | string } | { a1: string; formula: string } | { a1: string; format: string };

// Retired sidecar AJ:AM (FieldCap Total / Rotated / Rot Hrs / Rotating Hrs).
// Cleared on every refresh so they cannot come back. K = J-I, O = J-I-L.
const HELPER_COLS = ["AJ", "AK", "AL", "AM"] as const;

export interface DayRollContext {
  row: number;
  period: string;
  bha: number;
  priorBit: number;
  priorCirc: number;
  slideHours: number;
  startDepth: number | null;
}

export interface DayRollInputs {
  endDepth: string;
  bitTotal: string;
  circTotal: string;
}

export type DayRollCheck =
  | { ok: true; endDepth: number; bitDelta: number; circDelta: number }
  | { ok: false; error: string; bitDelta?: number; circDelta?: number };

type Raw = { v?: unknown; s?: unknown; f?: unknown } | null | undefined;

/** Data sheet reads with the writes made so far laid over the snapshot. */
class DayTable {
  private readonly writes = new Map<string, number | string>();
  private readonly formulas = new Map<string, string>();
  private readonly formats = new Map<string, string>();

  constructor(private readonly book: SnapshotBook) {}

  value(a1: string, sheet = "Data"): unknown {
    if (sheet === "Data" && this.writes.has(a1)) return this.writes.get(a1);
    return this.cell(sheet, a1)?.v ?? null;
  }

  set(a1: string, value: number | string): void {
    this.writes.set(a1, value);
  }

  /** Range.Formula = f, only where the sheet holds something else. */
  setFormula(a1: string, f: string): void {
    const current = this.cell("Data", a1);
    const held = typeof current?.f === "string" ? current.f : current?.v == null ? "" : String(current.v);
    if (held !== f) this.formulas.set(a1, f);
  }

  /** Range.NumberFormat = pattern, only where the cell shows something else. */
  setFormat(a1: string, pattern: string): void {
    if (this.pattern(a1) !== pattern) this.formats.set(a1, pattern);
  }

  written(a1: string): boolean {
    return this.writes.has(a1);
  }

  /** Snapshot value or formula still occupies the cell (not an overlay we already cleared). */
  occupied(a1: string): boolean {
    if (this.writes.has(a1)) return !blank(this.writes.get(a1));
    const current = this.cell("Data", a1);
    if (typeof current?.f === "string" && current.f.trim() !== "") return true;
    return !blank(current?.v);
  }

  clear(a1: string): void {
    if (this.occupied(a1)) this.writes.set(a1, "");
  }

  pending(): DayWrite[] {
    return [
      ...[...this.formulas].map(([a1, formula]) => ({ a1, formula })),
      ...[...this.writes].map(([a1, value]) => ({ a1, value })),
      ...[...this.formats].map(([a1, format]) => ({ a1, format })),
    ];
  }

  textFormatted(a1: string): boolean {
    return this.pattern(a1)?.includes("@") ?? false;
  }

  private pattern(a1: string): string | undefined {
    const style = this.cell("Data", a1)?.s;
    const styles = (this.book as { styles?: Record<string, { n?: { pattern?: string } } | undefined> }).styles;
    const resolved = typeof style === "string" ? styles?.[style] : (style as { n?: { pattern?: string } } | undefined);
    return resolved?.n?.pattern;
  }

  sheetRows(sheet: string): Record<string, Record<string, Raw> | null> | null {
    return (sheetOf(this.book, sheet)?.cellData as Record<string, Record<string, Raw> | null> | undefined) ?? null;
  }

  private cell(sheet: string, a1: string): Raw {
    const match = /^([A-Z]+)(\d+)$/.exec(a1);
    if (!match?.[1] || !match[2]) return null;
    const rows = this.sheetRows(sheet);
    return rows?.[Number(match[2]) - 1]?.[columnIndex(match[1])] ?? null;
  }
}

export function isDayPeriod(value: unknown): boolean {
  return typeof value === "string" && (PERIODS as readonly string[]).includes(value.trim());
}

function isClosingPeriod(value: string): boolean {
  return value === "Midnight" || value === "End of Run";
}

function isError(value: unknown): boolean {
  return typeof value === "string" && ERROR_TEXT.test(value.trim());
}

function blank(value: unknown): boolean {
  return value == null || isError(value) || String(value).trim() === "";
}

/** VBA IsNumeric on a cell value: numbers and numeric text. */
function numeric(value: unknown): number | null {
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  if (typeof value === "boolean") return value ? -1 : 0;
  if (typeof value !== "string" || value.trim() === "" || isError(value)) return null;
  const parsed = Number(value.trim());
  return Number.isFinite(parsed) ? parsed : null;
}

function num(value: unknown): number {
  return numeric(value) ?? 0;
}

function period(t: DayTable, row: number): string {
  const value = t.value(`H${row}`);
  return value == null ? "" : String(value).trim();
}

function sumColumn(t: DayTable, column: string, from: number, to: number): number {
  let total = 0;
  for (let row = from; row <= to; row += 1) total += numeric(t.value(`${column}${row}`)) ?? 0;
  return total;
}

/** Hours + minutes/60 per row; column V when its formula result is current. */
function sumSlideHours(t: DayTable, from: number, to: number): number {
  let total = 0;
  for (let row = from; row <= to; row += 1) {
    const stale = t.written(`M${row}`) || t.written(`N${row}`);
    const v = stale ? null : numeric(t.value(`V${row}`));
    if (v != null) total += v;
    else if (numeric(t.value(`M${row}`)) != null || numeric(t.value(`N${row}`)) != null) total += num(t.value(`M${row}`)) + num(t.value(`N${row}`)) / 60;
  }
  return total;
}

function findSourceRow(t: DayTable, bha: number): number {
  for (let row = SRC_FIRST; row <= SRC_LAST; row += 1) {
    const value = numeric(t.value(`H${row}`));
    if (value != null && Math.round(value) === bha) return row;
  }
  return 0;
}

function lastSourceBha(t: DayTable): number {
  let last = 0;
  for (let row = SRC_FIRST; row <= SRC_LAST; row += 1) {
    const value = numeric(t.value(`H${row}`));
    if (value != null) last = Math.round(value);
  }
  return last;
}

/** H3 must name a BHA in the H36:H47 mirror; otherwise the last mirror BHA is written there. */
function resolveActiveBha(t: DayTable): number {
  const header = numeric(t.value("H3"));
  if (header != null && findSourceRow(t, Math.round(header)) > 0) return Math.round(header);
  const last = lastSourceBha(t);
  if (last > 0) t.set("H3", last);
  // Empty mirror (DD Tools not built): keep the BHA the DD picked.
  else if (header != null && header > 0) return Math.round(header);
  return last;
}

interface SlideDay {
  key: number;
  metres: number;
  hours: number;
  rotMetres: number;
  rotHours: number;
}

/** One report day as FieldCap logged it, less the 6am / 4pm rows above it in that day. */
export interface FieldCapDay {
  slideM: number;
  slideH: number;
  rotM: number;
  rotH: number;
  totalM: number;
}

/** FieldCap's own BHA slide totals (_OC_BHA). The active BHA's mirror row is rebuilt from the day rows. */
function fieldCapBhaTotals(t: DayTable, bha: number): { metres: number; hours: number } | null {
  const rows = t.sheetRows("_OC_BHA");
  const header = rows?.[0];
  if (!rows || !header) return null;
  let cBha = -1;
  let cM = -1;
  let cH = -1;
  for (const [index, cell] of Object.entries(header)) {
    const name = String(cell?.v ?? "").trim();
    if (name === "BHA #") cBha = Number(index);
    if (name === "BHA Mtrs Slid") cM = Number(index);
    if (name === "BHA Hrs Slid") cH = Number(index);
  }
  if (cBha < 0 || (cM < 0 && cH < 0)) return null;
  for (const [index, row] of Object.entries(rows)) {
    if (Number(index) === 0) continue;
    const owner = numeric(row?.[cBha]?.v);
    if (owner == null || Math.round(owner) !== bha) continue;
    return { metres: cM < 0 ? 0 : num(row?.[cM]?.v), hours: cH < 0 ? 0 : num(row?.[cH]?.v) };
  }
  return null;
}

function slideDays(t: DayTable, bha: number): SlideDay[] {
  const rows = t.sheetRows("_OC_Slide");
  if (!rows) return [];
  let last = 0;
  for (const [index, row] of Object.entries(rows)) if (!blank(row?.[2]?.v)) last = Math.max(last, Number(index));
  const days: SlideDay[] = [];
  for (let index = 1; index <= last; index += 1) {
    const row = rows[index];
    const owner = numeric(row?.[2]?.v);
    if (owner == null || Math.round(owner) !== bha) continue;
    const rawKey = row?.[1]?.v;
    const parsedDate = typeof rawKey === "string" ? Date.parse(rawKey) : Number.NaN;
    const key = numeric(rawKey) ?? (Number.isFinite(parsedDate) ? parsedDate : index + 1);
    const metres = num(row?.[3]?.v);
    const rotMetres = num(row?.[4]?.v);
    const hours = num(row?.[5]?.v);
    const rotHours = num(row?.[6]?.v);
    const existing = days.find((day) => day.key === key);
    if (existing) {
      existing.metres += metres;
      existing.rotMetres += rotMetres;
      existing.hours += hours;
      existing.rotHours += rotHours;
    } else {
      days.push({ key, metres, hours, rotMetres, rotHours });
    }
  }
  return fillMissingDays(days.sort((a, b) => a.key - b.key));
}

const DATE_SERIAL_MIN = 36526; // 2000-01-01
const DATE_SERIAL_MAX = 73050; // 2099-12-31
const DAY_MS = 86_400_000;
const MAX_FILL_DAYS = 400;

/**
 * FieldCap writes a zero row for a day the BHA was in the hole without drilling. An older by-day
 * file skips that day, so the dates between the BHA's first and last FieldCap day become zero
 * days here; otherwise every later day would close the table row above its own (FillMissingDays).
 */
function fillMissingDays(days: SlideDay[]): SlideDay[] {
  const first = days[0];
  const last = days[days.length - 1];
  if (!first || !last || days.length < 2) return days;
  const serial = first.key >= DATE_SERIAL_MIN && last.key <= DATE_SERIAL_MAX;
  const epochMs = first.key >= Date.UTC(2000, 0, 1) && last.key < Date.UTC(2100, 0, 1);
  if (!serial && !epochMs) return days;
  const dayOf = (key: number) => (serial ? Math.floor(key) : Math.round(key / DAY_MS));
  const start = dayOf(first.key);
  const span = dayOf(last.key) - start + 1;
  if (span <= days.length || span > MAX_FILL_DAYS) return days;
  const filled: SlideDay[] = Array.from({ length: span }, (_, i) => ({
    key: serial ? start + i : (start + i) * DAY_MS,
    metres: 0,
    hours: 0,
    rotMetres: 0,
    rotHours: 0,
  }));
  for (const day of days) {
    const slot = filled[dayOf(day.key) - start];
    if (!slot) continue;
    slot.metres += day.metres;
    slot.hours += day.hours;
    slot.rotMetres += day.rotMetres;
    slot.rotHours += day.rotHours;
  }
  return filled;
}

function remainder(value: number, digits: number): number {
  return excelRound(Math.max(0, value), digits);
}

/**
 * FieldCap's report day that row closes (OpenCap by-day CSV), less the 6am / 4pm rows inside that
 * day. FieldCap is the source of truth for slide metres/time (L, M:N). Total / rotated metres stay
 * End-Start and End-Start-slide (K, O). Time Rotating is always P-V.
 */
function daySlideForRow(t: DayTable, bha: number, row: number): FieldCapDay | null {
  if (bha <= 0 || !isClosingPeriod(period(t, row))) return null;
  const days = slideDays(t, bha);
  if (days.length === 0) return null;
  let dayIndex = 0;
  let firstOfDay = DAY_FIRST;
  for (let r = DAY_FIRST; r <= row; r += 1) {
    if (!isClosingPeriod(period(t, r))) continue;
    dayIndex += 1;
    if (r < row) firstOfDay = r + 1;
  }
  const day = days[dayIndex - 1];
  if (!day) return null;
  return {
    slideM: remainder(day.metres - sumColumn(t, "L", firstOfDay, row - 1), 2),
    slideH: remainder(day.hours - sumSlideHours(t, firstOfDay, row - 1), 4),
    rotM: remainder(day.rotMetres, 2),
    rotH: remainder(day.rotHours, 4),
    totalM: remainder(day.metres + day.rotMetres, 2),
  };
}

function clearHelperBand(t: DayTable): void {
  for (let row = 2; row <= DAY_TOTALS; row += 1) {
    for (const column of HELPER_COLS) t.clear(`${column}${row}`);
  }
}

/** Daily-table formulas: K = J-I, O = J-I-L, S = O/(P-V). Email C9/C10/C16/C17 read rows 23/24. */
function ensureFieldCapFormulas(t: DayTable): void {
  clearHelperBand(t);
  for (let r = DAY_FIRST; r <= DAY_LAST; r += 1) {
    t.setFormula(`K${r}`, `=IFERROR(IF(ISBLANK(J${r}),"",J${r}-I${r}),"")`);
    t.setFormula(`O${r}`, `=IFERROR(IF(ISBLANK(J${r}),"",J${r}-I${r}-L${r}),"")`);
    t.setFormula(`S${r}`, `=IFERROR(IF(ISBLANK(P${r}),"",ROUND(O${r}/(P${r}-V${r}),2)),"")`);
  }
  t.setFormula("S24", `=IFERROR(ROUND(O24/(P24-V24),2),"")`);
  // BHA slide time from V24: SUM(N) skips minutes typed as text and never carries past 60.
  t.setFormula("M24", `=IFERROR(INT(ROUND(V24*60,0)/60),"")`);
  t.setFormula("N24", `=IFERROR(MOD(ROUND(V24*60,0),60),"")`);
  t.setFormula("C9", `=IFERROR(IF(IFERROR(N(V23),0)=0,"0Hrs",ROUND(V23,2)&"Hrs / "&ROUND(T23,4)*100&"% / "&ROUND(R23,2)&"m/hr"),"")`);
  t.setFormula("C10", `=IFERROR(IF(U23=0,"0Hrs",ROUND(P23-V23,2)&"Hrs / "&ROUND(U23,4)*100&"% / "&ROUND(S23,2)&"m/hr"),"")`);
  t.setFormula("C16", `=IFERROR(IF(IFERROR(N(V24),0)=0,"0Hrs",ROUND(V24,2)&"Hrs / "&ROUND(T24,4)*100&"% / "&ROUND(R24,2)&"m/hr"),"")`);
  t.setFormula("C17", `=IFERROR(IF(U24=0,"0Hrs",ROUND(P24-V24,2)&"Hrs / "&ROUND(U24,4)*100&"% / "&ROUND(S24,2)&"m/hr"),"")`);
  t.setFormula("C7", `=IFERROR(ROUND(IF(ISERROR(L23),0,L23),0),"")`);
  t.setFormula("C8", `=IFERROR(ROUND(O23,2),"")`);
  t.setFormula("C14", `=IFERROR(ROUND(L24,0),"")`);
  t.setFormula("C15", `=IFERROR(ROUND(O24,2),"")`);
  t.setFormula("C18", `=IFERROR(IF(U24=0,"0Hrs",ROUND(K24/P24,2)&"m/hr"),"")`);
  t.setFormula("C19", `=IFERROR(P24,"")`);
  t.setFormula("C20", `=IFERROR(Q24,"")`);
  t.setFormat("C7", `0"m"`);
  t.setFormat("C14", `0"m"`);
  for (let r = DAY_FIRST; r <= DAY_TOTALS; r += 1) t.setFormat(`L${r}`, "0");
}

function clearFieldCapDay(t: DayTable, row: number): void {
  for (const column of HELPER_COLS) t.clear(`${column}${row}`);
}

function slideIntervals(t: DayTable): { from: unknown[]; to: unknown[] } {
  const from: unknown[] = [];
  const to: unknown[] = [];
  for (let row = SS_FIRST; row <= SS_LAST; row += 1) {
    from.push(t.value(`R${row}`, "Slidesheet"));
    to.push(t.value(`S${row}`, "Slidesheet"));
  }
  return { from, to };
}

function hasSlideIntervals(t: DayTable): boolean {
  const { from, to } = slideIntervals(t);
  return from.some((a, i) => {
    const lo = numeric(a);
    const hi = numeric(to[i]);
    return lo != null && hi != null && hi > lo;
  });
}

function depthWindow(t: DayTable, row: number): { lo: number; hi: number } | null {
  const lo = numeric(t.value(`I${row}`));
  const hi = numeric(t.value(`J${row}`));
  return lo != null && hi != null && hi > lo ? { lo, hi } : null;
}

/** Slide metres inside the row's depth window from Slidesheet R/S; hours follow the metres. */
function depthSlideForRow(t: DayTable, bha: number, row: number, mirrorM: number, mirrorH: number): { metres: number; hours: number } | null {
  if (!hasSlideIntervals(t)) return null;
  const window = depthWindow(t, row);
  if (!window) return null;
  const { from, to } = slideIntervals(t);
  const metres = excelRound(slideMetersBetween(window.lo, window.hi, from, to), 2);
  const days = slideDays(t, bha);
  const poolM = days.length > 0 ? days.reduce((sum, day) => sum + day.metres, 0) : mirrorM;
  const poolH = days.length > 0 ? days.reduce((sum, day) => sum + day.hours, 0) : mirrorH;
  const hours = poolM > 0.0001 && metres > 0 && poolH > 0 ? Math.max(0, excelRound(poolH * (metres / poolM), 4)) : 0;
  return { metres, hours };
}

function writeMinutes(t: DayTable, row: number, minutes: number): void {
  const a1 = `N${row}`;
  t.set(a1, t.textFormatted(a1) ? String(minutes).padStart(2, "0") : minutes);
}

function splitHours(decimal: number): { hours: number; minutes: number } {
  if (decimal < 0) return { hours: 0, minutes: 0 };
  const total = excelRound(decimal * 60, 0);
  return { hours: Math.floor(total / 60), minutes: total % 60 };
}

function fillSlideMetrics(t: DayTable, row: number, overwrite: boolean): boolean {
  const bha = resolveActiveBha(t);
  if (bha === 0) return false;
  const source = findSourceRow(t, bha);
  const day = daySlideForRow(t, bha, row);
  const fieldCap = fieldCapBhaTotals(t, bha);
  // A FieldCap day needs neither BHA total, so an empty DD Tools mirror cannot stop it.
  if (source === 0 && !fieldCap && !day) return false;
  const mirrorM = fieldCap ? fieldCap.metres : source > 0 ? num(t.value(`I${source}`)) : 0;
  const mirrorH = fieldCap ? fieldCap.hours : source > 0 ? num(t.value(`K${source}`)) : 0;
  let deltaM = excelRound(mirrorM - sumColumn(t, "L", DAY_FIRST, row - 1), 2);
  let deltaH = excelRound(mirrorH - sumSlideHours(t, DAY_FIRST, row - 1), 4);
  let force = overwrite;
  // FieldCap's day is the source of truth; only rows it does not cover use the Slidesheet intervals.
  const depth = day ? null : depthSlideForRow(t, bha, row, mirrorM, mirrorH);
  if (day) {
    deltaM = day.slideM;
    deltaH = day.slideH;
    force = true;
  } else if (depth) {
    deltaM = depth.metres;
    deltaH = depth.hours;
    force = true;
  }
  let wrote = false;
  if (row > DAY_FIRST && blank(t.value(`I${row}`))) {
    const previousEnd = t.value(`J${row - 1}`);
    if (!blank(previousEnd) && numeric(previousEnd) != null) {
      t.set(`I${row}`, numeric(previousEnd)!);
      wrote = true;
    }
  }
  if (deltaM >= 0 && (force || blank(t.value(`L${row}`)))) {
    t.set(`L${row}`, deltaM);
    wrote = true;
  }
  const hoursBlank = blank(t.value(`M${row}`)) && blank(t.value(`N${row}`));
  if (deltaH >= 0 && (day != null || deltaH > 0 || hoursBlank) && (force || hoursBlank)) {
    const split = splitHours(deltaH);
    t.set(`M${row}`, split.hours);
    writeMinutes(t, row, split.minutes);
    wrote = true;
  }
  return wrote;
}

function lastPeriodRow(t: DayTable): number {
  for (let row = DAY_LAST; row >= DAY_FIRST; row -= 1) if (isDayPeriod(period(t, row))) return row;
  return 0;
}

/** DayRoll_BackfillSlideMetrics: the active BHA (Data!H3, from the H36:H47 mirror) refills the day rows. */
export function backfillDayBreakdown(book: SnapshotBook): DayWrite[] {
  const t = new DayTable(book);
  backfillSlideMetrics(t);
  return t.pending();
}

/**
 * DayRoll_SlideHoursWarning, read after the refresh has recalculated: day rows that do not match
 * FieldCap's day, and rows FieldCap does not cover whose slide time is past their bit hours. "" = none.
 */
export function dayBreakdownWarning(book: SnapshotBook): string {
  const t = new DayTable(book);
  const bha = Math.round(numeric(t.value("H3")) ?? 0);
  const offFieldCap: string[] = [];
  const overBit: string[] = [];
  for (let row = DAY_FIRST; row <= DAY_LAST; row += 1) {
    if (!isDayPeriod(period(t, row))) continue;
    const slideH = num(t.value(`M${row}`)) + num(t.value(`N${row}`)) / 60;
    const day = daySlideForRow(t, bha, row);
    if (day) {
      const off = [
        Math.abs(num(t.value(`L${row}`)) - day.slideM) > 0.005 ? "slide m" : "",
        Math.abs(slideH - day.slideH) > 0.01 ? "slide time" : "",
      ].filter(Boolean);
      if (off.length > 0) {
        offFieldCap.push(`  Row ${row} (${period(t, row)}): ${off.join(", ")}  (FieldCap S ${day.slideM.toFixed(2)} m / ${day.slideH.toFixed(2)} h)`);
      }
      continue;
    }
    const bit = numeric(t.value(`P${row}`));
    if (bit != null && !blank(t.value(`P${row}`)) && slideH > bit + 0.01) {
      overBit.push(`  Row ${row} (${period(t, row)}): ${slideH.toFixed(2)} h sliding, ${bit.toFixed(2)} bit hours`);
    }
  }
  const parts: string[] = [];
  if (offFieldCap.length > 0) {
    parts.push(`These day rows do not match FieldCap's day:\n${offFieldCap.join("\n")}\n\nREFRESH writes FieldCap's numbers. If a row still differs, its cells are locked or hold text.`);
  }
  if (overBit.length > 0) {
    parts.push(`FieldCap has no day for these rows, and Time Sliding is more than the bit hours:\n${overBit.join("\n")}\n\nRotating time would print negative in the email. Fix the slide time (M:N) or the bit hours (P) on those rows.`);
  }
  return parts.join("\n\n");
}

function backfillSlideMetrics(t: DayTable): void {
  ensureFieldCapFormulas(t);
  const active = lastPeriodRow(t);
  const bha = resolveActiveBha(t);
  const depthOn = hasSlideIntervals(t);
  for (let row = DAY_FIRST; row <= DAY_LAST; row += 1) {
    if (!daySlideForRow(t, bha, row)) clearFieldCapDay(t, row);
  }
  for (let row = DAY_FIRST; row <= DAY_LAST; row += 1) {
    if (!isDayPeriod(period(t, row))) continue;
    const hasDay = daySlideForRow(t, bha, row) != null;
    const needs = blank(t.value(`L${row}`)) || (blank(t.value(`M${row}`)) && blank(t.value(`N${row}`)));
    if (depthOn || hasDay || row === active || needs) fillSlideMetrics(t, row, depthOn || row === active || hasDay);
  }
}

function context(t: DayTable, row: number): DayRollContext {
  const stale = t.written(`M${row}`) || t.written(`N${row}`);
  const v = stale ? null : numeric(t.value(`V${row}`));
  const slide = v ?? (numeric(t.value(`M${row}`)) != null || numeric(t.value(`N${row}`)) != null ? num(t.value(`M${row}`)) + num(t.value(`N${row}`)) / 60 : 0);
  return {
    row,
    period: period(t, row),
    bha: Math.round(numeric(t.value("H3")) ?? 0),
    priorBit: sumColumn(t, "P", DAY_FIRST, row - 1),
    priorCirc: sumColumn(t, "Q", DAY_FIRST, row - 1),
    slideHours: excelRound(slide, 4),
    startDepth: numeric(t.value(`I${row}`)),
  };
}

/** A period was picked in H: fill start depth and slide metres/time, and describe the form. */
export function startDayRoll(book: SnapshotBook, row: number): { writes: DayWrite[]; context: DayRollContext } | null {
  if (row < DAY_FIRST || row > DAY_LAST) return null;
  const t = new DayTable(book);
  if (!isDayPeriod(period(t, row))) return null;
  fillSlideMetrics(t, row, true);
  return { writes: t.pending(), context: context(t, row) };
}

/** Form validation, as frmDayHours shows it while typing and on Apply. */
export function checkDayRoll(context: DayRollContext, inputs: DayRollInputs): DayRollCheck {
  const endText = inputs.endDepth.trim();
  const bitText = inputs.bitTotal.trim();
  const circText = inputs.circTotal.trim();
  if (!endText) return { ok: false, error: "End Depth is required." };
  const endDepth = numeric(endText);
  if (endDepth == null) return { ok: false, error: "End Depth must be numeric." };
  if (context.startDepth != null && endDepth + 1e-7 < context.startDepth) {
    return { ok: false, error: `End Depth must be >= Start Depth (${formatNumber(context.startDepth, 2)}).` };
  }
  if (endDepth <= 0) return { ok: false, error: "End Depth must be a positive number." };
  if (!bitText || !circText) return { ok: false, error: "Both Bit Total and Circ Total are required." };
  const bitTotal = numeric(bitText);
  const circTotal = numeric(circText);
  if (bitTotal == null || circTotal == null) return { ok: false, error: "Bit Total and Circ Total must be numeric." };
  const bitDelta = excelRound(bitTotal - context.priorBit, 4);
  const circDelta = excelRound(circTotal - context.priorCirc, 4);
  if (bitDelta < 0 || circDelta < 0) return { ok: false, error: "Bit and Circ deltas cannot be negative (running totals must not go backwards).", bitDelta, circDelta };
  if (circDelta + 1e-7 < bitDelta) return { ok: false, error: "Circ Hours must be >= Bit Hours for this row.", bitDelta, circDelta };
  if (context.period.toLowerCase() === "midnight" && circDelta > 24 + 1e-7) return { ok: false, error: "Circ Hours cannot exceed 24 for a Midnight row.", bitDelta, circDelta };
  return { ok: true, endDepth, bitDelta, circDelta };
}

/** Apply: End Depth into J, hour deltas into P/Q, then slide metres re-placed across every row. */
export function applyDayRoll(book: SnapshotBook, context: DayRollContext, inputs: DayRollInputs): { writes: DayWrite[] } | { error: string } {
  const check = checkDayRoll(context, inputs);
  if (!check.ok) return { error: check.error };
  const t = new DayTable(book);
  const row = context.row;
  t.set(`J${row}`, check.endDepth);
  t.set(`P${row}`, check.bitDelta);
  t.set(`Q${row}`, check.circDelta);
  if (check.bitDelta === 0) {
    if (blank(t.value(`L${row}`))) t.set(`L${row}`, 0);
    if (blank(t.value(`M${row}`))) t.set(`M${row}`, 0);
    if (blank(t.value(`N${row}`))) writeMinutes(t, row, 0);
  }
  backfillSlideMetrics(t);
  return { writes: t.pending() };
}

export function formatNumber(value: number, digits: number): string {
  return String(Number(value.toFixed(digits)));
}
