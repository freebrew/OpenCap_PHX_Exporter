import { slideMetersBetween } from "../../dd-math/src/project";
import { columnIndex, readColumn, readNumber, sheetOf, type SnapshotBook } from "./cells";

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
export const DAY_PERIOD_COLUMN = "H";
const PERIODS = ["6am", "4pm", "Midnight", "End of Run"] as const;
const SRC_FIRST = 36;
const SRC_LAST = 47;
const SS_FIRST = 13;
const SS_LAST = 505;
const ERROR_TEXT = /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|SPILL!|CALC!|GETTING_DATA)$/;

export interface DayWrite {
  a1: string;
  value: number | string;
}

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

type Raw = { v?: unknown; s?: unknown } | null | undefined;

/** Data sheet reads with the writes made so far laid over the snapshot. */
class DayTable {
  private readonly writes = new Map<string, number | string>();

  constructor(private readonly book: SnapshotBook) {}

  value(a1: string, sheet = "Data"): unknown {
    if (sheet === "Data" && this.writes.has(a1)) return this.writes.get(a1);
    return this.cell(sheet, a1)?.v ?? null;
  }

  set(a1: string, value: number | string): void {
    this.writes.set(a1, value);
  }

  written(a1: string): boolean {
    return this.writes.has(a1);
  }

  pending(): DayWrite[] {
    return [...this.writes].map(([a1, value]) => ({ a1, value }));
  }

  textFormatted(a1: string): boolean {
    const style = this.cell("Data", a1)?.s;
    const styles = (this.book as { styles?: Record<string, { n?: { pattern?: string } } | undefined> }).styles;
    const resolved = typeof style === "string" ? styles?.[style] : (style as { n?: { pattern?: string } } | undefined);
    return resolved?.n?.pattern?.includes("@") ?? false;
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

/** VBA Round: banker's rounding. */
function vbaRound(value: number, digits: number): number {
  const factor = 10 ** digits;
  const scaled = value * factor;
  const floor = Math.floor(scaled);
  const diff = scaled - floor;
  const epsilon = 1e-9 * Math.max(1, Math.abs(scaled));
  let rounded: number;
  if (Math.abs(diff - 0.5) < epsilon) rounded = floor % 2 === 0 ? floor : floor + 1;
  else rounded = Math.round(scaled);
  return rounded / factor;
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
  return last;
}

interface SlideDay {
  key: number;
  metres: number;
  hours: number;
  rotHours: number;
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
    const hours = num(row?.[5]?.v);
    const rotHours = num(row?.[6]?.v);
    const existing = days.find((day) => day.key === key);
    if (existing) {
      existing.metres += metres;
      existing.hours += hours;
      existing.rotHours += rotHours;
    } else {
      days.push({ key, metres, hours, rotHours });
    }
  }
  return days.sort((a, b) => a.key - b.key);
}

/**
 * OpenCap by-day slide for the report day that row closes, less the 6am / 4pm rows inside that day.
 * FieldCap logs a slide stand as one Sliding block, connections and surveys included, so its
 * Sliding + Drilling time runs past Pason's on-bottom hours. The day keeps FieldCap's
 * slide-to-rotate split of the bit hours instead, and slide time never exceeds bit time.
 */
function daySlideForRow(t: DayTable, bha: number, row: number): { metres: number; hours: number } | null {
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
  const bitDay = sumColumn(t, "P", firstOfDay, row);
  const dayHours = bitDay > 0 && day.hours + day.rotHours > bitDay ? (day.hours * bitDay) / (day.hours + day.rotHours) : day.hours;
  const metres = Math.max(0, day.metres - sumColumn(t, "L", firstOfDay, row - 1));
  const hours = Math.max(0, dayHours - sumSlideHours(t, firstOfDay, row - 1));
  return { metres: vbaRound(metres, 2), hours: vbaRound(hours, 4) };
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

/** The row ends deeper than the last Slide To on the Slidesheet. */
function rowEndsPastSlides(t: DayTable, row: number): boolean {
  const window = depthWindow(t, row);
  if (!window) return false;
  const { from, to } = slideIntervals(t);
  let deepest = 0;
  from.forEach((a, i) => {
    const lo = numeric(a);
    const hi = numeric(to[i]);
    if (lo != null && hi != null && hi > lo && hi > deepest) deepest = hi;
  });
  return window.hi > deepest + 0.5;
}

/** Slide metres inside the row's depth window from Slidesheet R/S; hours follow the metres. */
function depthSlideForRow(t: DayTable, bha: number, row: number, mirrorM: number, mirrorH: number): { metres: number; hours: number } | null {
  if (!hasSlideIntervals(t)) return null;
  const window = depthWindow(t, row);
  if (!window) return null;
  const { from, to } = slideIntervals(t);
  const metres = vbaRound(slideMetersBetween(window.lo, window.hi, from, to), 2);
  const days = slideDays(t, bha);
  const poolM = days.length > 0 ? days.reduce((sum, day) => sum + day.metres, 0) : mirrorM;
  const poolH = days.length > 0 ? days.reduce((sum, day) => sum + day.hours, 0) : mirrorH;
  const hours = poolM > 0.0001 && metres > 0 && poolH > 0 ? Math.max(0, vbaRound(poolH * (metres / poolM), 4)) : 0;
  return { metres, hours };
}

function writeMinutes(t: DayTable, row: number, minutes: number): void {
  const a1 = `N${row}`;
  t.set(a1, t.textFormatted(a1) ? String(minutes).padStart(2, "0") : minutes);
}

function splitHours(decimal: number): { hours: number; minutes: number } {
  if (decimal < 0) return { hours: 0, minutes: 0 };
  const total = vbaRound(decimal * 60, 0);
  return { hours: Math.floor(total / 60), minutes: total % 60 };
}

function fillSlideMetrics(t: DayTable, row: number, overwrite: boolean): boolean {
  const bha = resolveActiveBha(t);
  if (bha === 0) return false;
  const source = findSourceRow(t, bha);
  if (source === 0) return false;
  const fieldCap = fieldCapBhaTotals(t, bha);
  const mirrorM = fieldCap ? fieldCap.metres : num(t.value(`I${source}`));
  const mirrorH = fieldCap ? fieldCap.hours : num(t.value(`K${source}`));
  let deltaM = vbaRound(mirrorM - sumColumn(t, "L", DAY_FIRST, row - 1), 2);
  let deltaH = vbaRound(mirrorH - sumSlideHours(t, DAY_FIRST, row - 1), 4);
  let force = overwrite;
  const day = daySlideForRow(t, bha, row);
  if (day) {
    deltaM = day.metres;
    deltaH = day.hours;
    force = true;
  }
  // The Slidesheet gets a slide when the next survey goes in, so a row that ends past its last
  // slide can be missing the newest ones. FieldCap's day already has them.
  const depth = depthSlideForRow(t, bha, row, mirrorM, mirrorH);
  if (depth) {
    deltaM = day && day.metres > depth.metres + 0.005 && rowEndsPastSlides(t, row) ? day.metres : depth.metres;
    if (!day) deltaH = depth.hours;
    force = true;
  }
  // The depth share and the BHA delta know nothing of this row's bit hours.
  const rowBit = num(t.value(`P${row}`));
  if (rowBit > 0 && deltaH > rowBit) deltaH = rowBit;
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
  if (deltaH >= 0 && (deltaH > 0 || hoursBlank) && (force || hoursBlank)) {
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

function backfillSlideMetrics(t: DayTable): void {
  const active = lastPeriodRow(t);
  const bha = resolveActiveBha(t);
  const depthOn = hasSlideIntervals(t);
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
    slideHours: vbaRound(slide, 4),
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
  const bitDelta = vbaRound(bitTotal - context.priorBit, 4);
  const circDelta = vbaRound(circTotal - context.priorCirc, 4);
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
