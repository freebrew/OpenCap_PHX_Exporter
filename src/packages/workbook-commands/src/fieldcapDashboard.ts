import { looksLikeSerial, motorLike } from "../../strapsheets-core/src/toolKinds";
import { type SurveyStation } from "../../fieldcap-sync/src/surveys";
import { serialHours, writeSerialHoursTable } from "./bhaSelect";
import { bhaNumber, numOrZero, OcTable } from "./ocTable";
import { GOOD_SURVEY_FILL, highlightNewestSurvey, isGoodSurveyRow, lastSurveyRow } from "./slidesheet";
import { excelRound, fixed2, numeric, type CellScalar, type SheetBook } from "./sheetBook";
import { toolHoursSync } from "./toolHours";

// FieldCap's BHA pull lands on _OC_BHA. Data!H35:M47 are formulas that read the
// DD Tools raw summary (O5:U17), and Data!H3 names the BHA whose totals fill the
// day rows. This writes that summary and the active motor (Data!C22). It does
// not replace the Data formulas.

const DATA = "Data";
const TOOLS = "DD Tools";
const SS = "Slidesheet";
const MIRROR_FIRST = 6;
const MIRROR_LAST = 17;
const SS_FIRST = 13;
const SS_LAST = 505;
/** The tie-in row: its column B holds the Survey Offset every row below carries down. */
const SS_TIE_IN = 12;
const COL_B = 2;
const COL_D = 4;
const COL_E = 5;
const COL_F = 6;
const COL_G = 7;
const COL_H = 8;
const SURVEY_DEPTH_TOLERANCE_M = 0.05;
const SURVEY_TVD_CONFIRM_M = 0.1;
/** Two survey offsets closer than this are the same offset. */
const SURVEY_OFFSET_SAME_M = 0.005;

const HEADERS = ["BHA", "Meters", "Mtrs Sld", "Mtrs Rot", "Hrs Sld", "Hrs Rot", "Hrs Crc", "Total Hrs", "Below Rot"] as const;

export interface FieldCapDashboard {
  bhas: number;
  activeBha: number;
  motor: string;
  tools: number;
  motors: number;
  serials: number;
}

export type SurveyDepthCoordinate = "bit-depth" | "survey-depth";

export interface SurveyTvdMismatch {
  row: number;
  md: number;
  slidesheetTvd: number;
  fieldCapTvd: number;
  difference: number;
}

export interface SurveySyncOptions {
  allowTvdMismatch?: boolean;
}

/** A Slidesheet bit depth that was replaced because FieldCap's survey is the authority. */
export interface SurveyDepthCorrection {
  row: number;
  md: number;
  from: number;
  to: number;
}

/** A row where the Survey Offset became what FieldCap's bit depth minus survey depth says. */
export interface SurveyOffsetChange {
  row: number;
  md: number;
  /** The offset the row carried before; null when the sheet had none. */
  from: number | null;
  to: number;
}

export interface SurveySyncResult {
  coordinate: SurveyDepthCoordinate | null;
  matched: number;
  changed: number;
  unchanged: number;
  appended: number;
  ignored: number;
  skippedPlaceholders: number;
  needsConfirmation: boolean;
  tvdMismatches: SurveyTvdMismatch[];
  /** Bit depths the Slidesheet had that FieldCap's survey replaced. */
  depthCorrections: SurveyDepthCorrection[];
  rejectedMd: number | null;
  warning: string | null;
  /** Rows where FieldCap's bit depth put a new Survey Offset (column B), and the offset they replaced. */
  offsetChanges: SurveyOffsetChange[];
  /** What was done when an empty Slidesheet was filled from row 13; null otherwise. */
  note: string | null;
  /** True when that fill had to use the sheet's own Survey Offset because FieldCap gave no bit depths. */
  checkOffset: boolean;
}

interface MirrorRow {
  bha: number;
  status: string;
  motor: string;
  metres: number;
  slid: number;
  rot: number;
  hrsSlid: number;
  hrsRot: number;
  hrsCirc: number;
  hrsTotal: number;
  hrsBelow: number;
  components: Array<{ serial: string; code: string; description: string }>;
}

function blank(value: unknown): boolean {
  return value == null || String(value).trim() === "";
}

function rowsOf(book: SheetBook): MirrorRow[] {
  const table = new OcTable(book, "_OC_BHA");
  if (!table.exists) return [];
  const cBha = table.col("BHA #");
  if (cBha === 0) return [];
  const cStatus = table.col("Status");
  const cMotor = table.col("Motor");
  const cMetres = table.col("Metres Drilled");
  const cSlid = table.col("BHA Mtrs Slid");
  const cRot = table.col("BHA Mtrs Rot");
  const cHrsSlid = table.col("BHA Hrs Slid");
  const cHrsRot = table.col("BHA Hrs Rot");
  const cHrsCirc = table.col("BHA Hrs Circ");
  const cHrsTotal = table.col("BHA Total Hrs");
  const cHrsBelow = table.col("BHA Below Rot");
  const cSerial = table.col("Serial #");
  const cCode = table.col("Item Code");
  const cDesc = table.col("Description");
  const grouped = new Map<number, MirrorRow>();
  const last = table.lastRow(cBha);
  for (let row = 2; row <= last; row += 1) {
    const bha = bhaNumber(table.value(row, cBha));
    if (bha <= 0) continue;
    let record = grouped.get(bha);
    if (!record) {
      record = {
        bha,
        status: cStatus ? table.str(row, cStatus).trim() : "",
        motor: cMotor ? table.str(row, cMotor).trim() : "",
        metres: cMetres ? numOrZero(table.value(row, cMetres)) : 0,
        slid: cSlid ? numOrZero(table.value(row, cSlid)) : 0,
        rot: cRot ? numOrZero(table.value(row, cRot)) : 0,
        hrsSlid: cHrsSlid ? numOrZero(table.value(row, cHrsSlid)) : 0,
        hrsRot: cHrsRot ? numOrZero(table.value(row, cHrsRot)) : 0,
        hrsCirc: cHrsCirc ? numOrZero(table.value(row, cHrsCirc)) : 0,
        hrsTotal: cHrsTotal ? numOrZero(table.value(row, cHrsTotal)) : 0,
        hrsBelow: cHrsBelow ? numOrZero(table.value(row, cHrsBelow)) : 0,
        components: [],
      };
      grouped.set(bha, record);
    }
    const serial = cSerial ? table.str(row, cSerial).trim() : "";
    const code = cCode ? table.str(row, cCode).trim() : "";
    const description = cDesc ? table.str(row, cDesc).trim() : "";
    if (serial !== "" || code !== "" || description !== "") record.components.push({ serial, code, description });
    if (record.motor === "" && cMotor) record.motor = table.str(row, cMotor).trim();
  }
  return [...grouped.values()].sort((a, b) => a.bha - b.bha).slice(0, MIRROR_LAST - MIRROR_FIRST + 1);
}

export interface DashboardOptions {
  /** Honor the BHA currently typed on Data!H3 when it exists in the mirror (manual BHA switch). */
  preserveSelection?: boolean;
  /** Never downgrade Data!H3 to a lower BHA # than it already holds (protects against stale CSVs). */
  noDowngrade?: boolean;
}

function activeBha(rows: MirrorRow[], current: number, options?: DashboardOptions): number {
  if (options?.preserveSelection && rows.some((row) => row.bha === current)) return current;
  const running = [...rows].reverse().find((row) => /active/i.test(row.status) && !/complete/i.test(row.status));
  const candidate = running?.bha ?? rows[rows.length - 1]?.bha ?? 0;
  if (options?.noDowngrade && current > 0 && candidate < current) return current;
  return candidate;
}

function motorSerial(row: MirrorRow | undefined): string {
  if (!row) return "";
  if (looksLikeSerial(row.motor)) return row.motor;
  const component = row.components.find((entry) => motorLike(entry.serial, entry.description, entry.code));
  return component?.serial ?? "";
}

function writeMirror(book: SheetBook, rows: MirrorRow[]): void {
  if (!book.hasSheet(TOOLS)) return;
  if (blank(book.value(TOOLS, 3, 15))) book.setValue(TOOLS, 3, 15, "RAW BHA SUMMARY");
  HEADERS.forEach((header, index) => {
    if (blank(book.value(TOOLS, 5, 15 + index))) book.setValue(TOOLS, 5, 15 + index, header);
  });
  for (let slot = 0; slot < MIRROR_LAST - MIRROR_FIRST + 1; slot += 1) {
    const row = rows[slot];
    const line: CellScalar[] = row
      ? [row.bha, row.metres, row.slid, row.rot, row.hrsSlid, row.hrsRot, row.hrsCirc, row.hrsTotal, row.hrsBelow]
      : [null, null, null, null, null, null, null, null, null];
    line.forEach((value, index) => book.setValue(TOOLS, MIRROR_FIRST + slot, 15 + index, value));
  }
}

function writeComponents(book: SheetBook, row: MirrorRow | undefined): void {
  if (!book.hasSheet(TOOLS)) return;
  const headers = ["#", "SERIAL #", "ITEM CODE", "DESCRIPTION", "HRS SLD", "HRS ROT", "TOTAL HRS", "BHAs"];
  headers.forEach((header, index) => book.setValue(TOOLS, 5, 1 + index, header));
  const components = row?.components ?? [];
  for (let index = 0; index < 36; index += 1) {
    const component = components[index];
    const line: CellScalar[] = component
      ? [index + 1, component.serial, component.code, component.description, row?.hrsSlid ?? 0, row?.hrsRot ?? 0, row?.hrsTotal ?? 0, row?.bha ?? null]
      : [null, null, null, null, null, null, null, null];
    line.forEach((value, col) => book.setValue(TOOLS, 6 + index, 1 + col, value));
  }
}

/**
 * Project the loaded FieldCap BHA onto DD Tools, Data!H3 and Data!C22, then the
 * serial-hours tables. Data!H35:M47 stay formulas. Returns null when Data is missing.
 */
export function projectFieldCapDashboard(book: SheetBook, options?: DashboardOptions): FieldCapDashboard | null {
  if (!book.hasSheet(DATA)) return null;
  const rows = rowsOf(book);
  writeMirror(book, rows);
  const header = bhaNumber(book.value(DATA, 3, 8));
  const active = activeBha(rows, header, options);
  if (active > 0 && !book.hasFormula(DATA, 3, 8) && header !== active) book.setValue(DATA, 3, 8, active);
  const chosen = rows.find((row) => row.bha === (active || header));
  const motor = motorSerial(chosen);
  if (motor !== "" && !book.hasFormula(DATA, 22, 3)) book.setValue(DATA, 22, 3, motor);
  writeComponents(book, chosen);
  const hours = toolHoursSync(book);
  const tools = writeSerialHoursTable(book, serialHours(book));
  return {
    bhas: rows.length,
    activeBha: active,
    motor,
    tools: hours?.tools ?? 0,
    motors: hours?.motors ?? 0,
    serials: tools,
  };
}

interface SlidesheetSurveyRow {
  row: number;
  d: number | null;
  e: number | null;
  good: boolean;
}

interface MatchedSurvey {
  station: SurveyStation;
  target: SlidesheetSurveyRow;
  /** The bit depth to write when the row's own depth is not FieldCap's. Absent when it already agrees. */
  depth?: number;
}

function emptySurveySyncResult(): SurveySyncResult {
  return {
    coordinate: null,
    matched: 0,
    changed: 0,
    unchanged: 0,
    appended: 0,
    ignored: 0,
    skippedPlaceholders: 0,
    needsConfirmation: false,
    tvdMismatches: [],
    depthCorrections: [],
    offsetChanges: [],
    rejectedMd: null,
    warning: null,
    note: null,
    checkOffset: false,
  };
}

function closeSurveyDepth(left: number, right: number): boolean {
  return Math.abs(left - right) <= SURVEY_DEPTH_TOLERANCE_M + 1e-9;
}

function slidesheetSurveyRows(book: SheetBook): SlidesheetSurveyRow[] {
  const rows: SlidesheetSurveyRow[] = [];
  for (let row = SS_FIRST; row <= SS_LAST; row += 1) {
    const d = numeric(book.value(SS, row, COL_D));
    const e = numeric(book.value(SS, row, COL_E));
    if (d == null && e == null) continue;
    rows.push({ row, d, e, good: isGoodSurveyRow(book, row) });
  }
  return rows;
}

function depthAt(row: SlidesheetSurveyRow, coordinate: SurveyDepthCoordinate): number | null {
  return coordinate === "bit-depth" ? row.d : row.e;
}

function coordinateScore(rows: SlidesheetSurveyRow[], stations: readonly SurveyStation[], coordinate: SurveyDepthCoordinate): number {
  const goodDepths = rows.filter((row) => row.good).map((row) => depthAt(row, coordinate)).filter((depth): depth is number => depth != null);
  return stations.reduce((score, station) => score + (goodDepths.some((depth) => closeSurveyDepth(depth, station.md)) ? 1 : 0), 0);
}

function chooseSurveyCoordinate(rows: SlidesheetSurveyRow[], stations: readonly SurveyStation[]): SurveyDepthCoordinate | null {
  const bitScore = coordinateScore(rows, stations, "bit-depth");
  const surveyScore = coordinateScore(rows, stations, "survey-depth");
  if (bitScore === 0 && surveyScore === 0) return null;
  if (bitScore !== surveyScore) return bitScore > surveyScore ? "bit-depth" : "survey-depth";
  const goodRows = rows.filter((row) => row.good);
  return goodRows.every((row) => row.d != null && row.e != null && closeSurveyDepth(row.d, row.e)) ? "bit-depth" : null;
}

function rejectedSurveyResult(
  coordinate: SurveyDepthCoordinate | null,
  rejectedMd: number | null,
  reason: string,
  ignored = 0,
  skippedPlaceholders = 0,
): SurveySyncResult {
  const depth = rejectedMd == null ? "" : ` at MD ${fixed2(rejectedMd)}`;
  return {
    ...emptySurveySyncResult(),
    coordinate,
    ignored,
    skippedPlaceholders,
    rejectedMd,
    warning: `FieldCap surveys were not changed${depth}: ${reason}`,
  };
}

function withoutTrailingPlaceholders(stations: readonly SurveyStation[]): { stations: SurveyStation[]; skipped: number } {
  const kept = [...stations].sort((left, right) => left.md - right.md);
  let skipped = 0;
  while (kept.length > 0) {
    const last = kept[kept.length - 1]!;
    if (last.inc !== 0 || last.azm !== 0) break;
    kept.pop();
    skipped += 1;
  }
  return { stations: kept, skipped };
}

function surveyOffset(rows: SlidesheetSurveyRow[]): number | null {
  for (let index = rows.length - 1; index >= 0; index -= 1) {
    const row = rows[index]!;
    if (row.d != null && row.e != null) return row.d - row.e;
  }
  return null;
}

function tvdCorrectionMismatches(book: SheetBook, matched: MatchedSurvey[], lastGoodRow: number): SurveyTvdMismatch[] | null {
  const mismatches: SurveyTvdMismatch[] = [];
  for (const { station, target, depth } of matched) {
    if (target.row > lastGoodRow) continue;
    const changesDirection = numeric(book.value(SS, target.row, COL_F)) !== station.inc || numeric(book.value(SS, target.row, COL_G)) !== station.azm;
    // A depth that follows FieldCap moves the survey's TVD, so it goes through the same check.
    if (!changesDirection && depth === undefined) continue;
    const slidesheetTvd = numeric(book.value(SS, target.row, COL_H));
    if (station.tvd == null || slidesheetTvd == null) {
      // FieldCap is the authority on a depth; only a direction change has to be comparable.
      if (changesDirection) return null;
      continue;
    }
    const difference = Math.abs(slidesheetTvd - station.tvd);
    if (difference > SURVEY_TVD_CONFIRM_M + 1e-9) {
      mismatches.push({
        row: target.row,
        md: station.md,
        slidesheetTvd,
        fieldCapTvd: station.tvd,
        difference,
      });
    }
  }
  return mismatches;
}

function offsetText(offset: number): string {
  return String(excelRound(offset, 2));
}

/**
 * A Slidesheet with no depth rows yet (a new well) has nothing to reconcile against, so every
 * FieldCap station below the tie-in goes on from row 13. The Survey Offset comes from FieldCap
 * (bit depth minus survey depth) and is written to the tie-in row, and again on the row where
 * it changes. Without FieldCap bit depths the sheet's own offset is used and the result says so.
 * Bit Depth is survey depth plus that offset, so the sheet's survey depth equals FieldCap's.
 */
function seedEmptySlidesheet(book: SheetBook, stations: readonly SurveyStation[], skippedPlaceholders: number): SurveySyncResult {
  const tieIn = numeric(book.value(SS, SS_TIE_IN, COL_E));
  const kept = tieIn == null ? [...stations] : stations.filter((station) => station.md > tieIn + SURVEY_DEPTH_TOLERANCE_M);
  const ignored = stations.length - kept.length;
  const first = kept[0];
  if (!first) return { ...emptySurveySyncResult(), ignored, skippedPlaceholders };
  const lastRow = SS_FIRST + kept.length - 1;
  if (lastRow > SS_LAST) {
    return rejectedSurveyResult(null, kept[SS_LAST - SS_FIRST + 1]?.md ?? first.md, `would exceed Slidesheet row ${SS_LAST}.`, ignored, skippedPlaceholders);
  }
  for (let row = SS_FIRST; row <= lastRow; row += 1) {
    if (book.hasFormula(SS, row, COL_D)) return rejectedSurveyResult(null, kept[row - SS_FIRST]?.md ?? first.md, `cannot use Slidesheet row ${row}.`, ignored, skippedPlaceholders);
  }

  const sheetOffset = numeric(book.value(SS, SS_TIE_IN, COL_B));
  const fieldCapOffsets = kept
    .map((station) => (station.bit != null && station.bit >= station.md - SURVEY_DEPTH_TOLERANCE_M ? excelRound(station.bit - station.md, 2) : null))
    .filter((offset): offset is number => offset != null);
  const fromFieldCap = fieldCapOffsets.length === kept.length;
  if (!fromFieldCap && sheetOffset == null) {
    return rejectedSurveyResult(null, first.md, `FieldCap gave no bit depths and the Slidesheet has no Survey Offset on row ${SS_TIE_IN}.`, ignored, skippedPlaceholders);
  }

  const notes: string[] = [`The Slidesheet had no surveys, so it was filled from row ${SS_FIRST}.`];
  let running = sheetOffset;
  kept.forEach((station, index) => {
    const row = SS_FIRST + index;
    const offset = fromFieldCap ? fieldCapOffsets[index] ?? 0 : sheetOffset ?? 0;
    if (running == null || Math.abs(offset - running) > SURVEY_OFFSET_SAME_M) {
      if (index === 0) {
        book.setValue(SS, SS_TIE_IN, COL_B, offset);
        notes.push(`Survey Offset was set to ${offsetText(offset)} from FieldCap${sheetOffset == null ? "" : ` (it was ${offsetText(sheetOffset)})`}.`);
      } else {
        book.setValue(SS, row, COL_B, offset);
        notes.push(`FieldCap's offset changes to ${offsetText(offset)} at row ${row}.`);
      }
      running = offset;
    } else if (index === 0 && fromFieldCap) notes.push(`Survey Offset is ${offsetText(offset)}, as in FieldCap.`);
    book.setValue(SS, row, COL_D, excelRound(station.md + offset, 2));
    book.setValue(SS, row, COL_F, station.inc);
    book.setValue(SS, row, COL_G, station.azm);
    book.setFill(SS, { row, col: COL_F, row2: row, col2: COL_G }, GOOD_SURVEY_FILL);
  });
  if (!fromFieldCap && sheetOffset != null) {
    notes.push(`FieldCap gave no bit depths, so Bit Depth uses the Survey Offset on the sheet (${offsetText(sheetOffset)}). Check that offset.`);
  }
  highlightNewestSurvey(book, lastRow);

  return {
    ...emptySurveySyncResult(),
    coordinate: "survey-depth",
    changed: kept.length,
    appended: kept.length,
    ignored,
    skippedPlaceholders,
    note: notes.join(" "),
    checkOffset: !fromFieldCap,
  };
}

/** The bit depth FieldCap's survey puts on a row: its depth in the sheet's coordinate, plus the row's survey offset when that is the coordinate. */
function depthForStation(station: SurveyStation, coordinate: SurveyDepthCoordinate, offset: number | null): number | null {
  if (coordinate === "bit-depth") return excelRound(station.md, 2);
  // FieldCap's bit depth for the survey is the authority; the row's offset only stands in when it gives none.
  if (station.bit != null) return excelRound(station.bit, 2);
  return offset == null ? null : excelRound(station.md + offset, 2);
}

/**
 * The Survey Offset (column B) follows FieldCap's bit depth. Each BHA has its own distance from the
 * bit to the survey sensor, so the offset changes where a BHA's surveys begin, and FieldCap says so
 * only through each survey's bit depth. Walking the rows from the tie-in, a written survey whose
 * bit depth minus survey depth is not the offset the row carries gets that offset typed into B
 * (the rows below carry it down through their own formula) and its Bit Depth set to FieldCap's.
 * Survey depth (D minus B) is unchanged, so the survey stays where FieldCap has it.
 */
function followFieldCapOffsets(
  book: SheetBook,
  written: ReadonlyArray<{ row: number; station: SurveyStation }>,
  depthCorrections: SurveyDepthCorrection[],
): SurveyOffsetChange[] {
  const changes: SurveyOffsetChange[] = [];
  const byRow = new Map(written.filter(({ station }) => station.bit != null).map((entry) => [entry.row, entry.station]));
  if (byRow.size === 0) return changes;
  const lastRow = Math.max(...byRow.keys());
  let chain = numeric(book.value(SS, SS_TIE_IN, COL_B));
  for (let row = SS_FIRST; row <= lastRow; row += 1) {
    const typed = book.hasFormula(SS, row, COL_B) ? null : numeric(book.value(SS, row, COL_B));
    let carried = typed ?? chain;
    const station = byRow.get(row);
    if (station?.bit != null) {
      const target = excelRound(station.bit - station.md, 2);
      if (carried == null || Math.abs(carried - target) > SURVEY_DEPTH_TOLERANCE_M + 1e-9) {
        book.setValue(SS, row, COL_B, target);
        changes.push({ row, md: station.md, from: carried, to: target });
        carried = target;
      }
      const bit = excelRound(station.bit, 2);
      const had = numeric(book.value(SS, row, COL_D));
      if (had !== bit) {
        book.setValue(SS, row, COL_D, bit);
        if (had != null && !depthCorrections.some((fix) => fix.row === row)) depthCorrections.push({ row, md: station.md, from: had, to: bit });
      }
    }
    chain = carried;
  }
  return changes;
}

function ownOffset(row: SlidesheetSurveyRow | undefined): number | null {
  return row && row.d != null && row.e != null ? row.d - row.e : null;
}

/**
 * Reconcile FieldCap's complete history with the Slidesheet. FieldCap is the authority.
 *
 * Surveys are put on rows in order, not only by depth. A station that matches a good survey row's
 * depth (within 0.05 m) anchors that row. A good row between anchors whose depth FieldCap does not
 * share takes the station FieldCap has in that place, and its bit depth is replaced. Every station
 * deeper than the last good row goes on the next rows in order, on top of the planned depths there.
 * Rows are never inserted, and history is never replayed into blank rows.
 * A Slidesheet with no depth rows at all is filled from row 13 instead.
 */
export function writeSurveyStations(book: SheetBook, stations: readonly SurveyStation[], options: SurveySyncOptions = {}): SurveySyncResult {
  const empty = emptySurveySyncResult();
  const prepared = withoutTrailingPlaceholders(stations);
  empty.skippedPlaceholders = prepared.skipped;
  if (!book.hasSheet(SS) || prepared.stations.length === 0) return empty;
  const rows = slidesheetSurveyRows(book);
  if (rows.length === 0) return seedEmptySlidesheet(book, prepared.stations, prepared.skipped);
  const goodRows = rows.filter((row) => row.good);
  if (goodRows.length === 0) {
    return rejectedSurveyResult(null, prepared.stations[0]?.md ?? null, "no existing good survey establishes the depth coordinate.", 0, prepared.skipped);
  }
  const coordinate = chooseSurveyCoordinate(rows, prepared.stations);
  if (!coordinate) {
    return rejectedSurveyResult(null, prepared.stations[0]?.md ?? null, "bit depth and survey depth could not be distinguished.", 0, prepared.skipped);
  }

  const list = prepared.stations;
  const skipped = prepared.skipped;
  let ignored = 0;
  const reject = (md: number | null, reason: string) => rejectedSurveyResult(coordinate, md, reason, ignored, skipped);
  const byRow = new Map(rows.map((row) => [row.row, row]));
  const lastGoodRow = goodRows[goodRows.length - 1]!.row;
  const sheetOffset = coordinate === "survey-depth" ? surveyOffset(rows) : 0;

  // 1. Anchors: stations that share a good survey row's depth.
  const rowOf: Array<SlidesheetSurveyRow | undefined> = list.map(() => undefined);
  const anchors: Array<{ index: number; row: SlidesheetSurveyRow }> = [];
  for (let index = 0; index < list.length; index += 1) {
    const station = list[index]!;
    const targets = goodRows.filter((row) => {
      const depth = depthAt(row, coordinate);
      return depth != null && closeSurveyDepth(depth, station.md);
    });
    if (targets.length === 0) continue;
    if (targets.length > 1) return reject(station.md, "matches more than one Slidesheet row.");
    const target = targets[0]!;
    const previous = anchors[anchors.length - 1];
    if (previous && target.row <= previous.row.row) return reject(station.md, `matches Slidesheet row ${target.row}, which is out of order.`);
    rowOf[index] = target;
    anchors.push({ index, row: target });
  }
  if (anchors.length === 0) return reject(list[0]?.md ?? null, "no FieldCap survey matches a good Slidesheet survey.");

  // 2. Good rows no station shares: FieldCap's station in that place replaces their depth.
  const pairs: Array<{ index: number; row: SlidesheetSurveyRow }> = [];
  const goodBetween = (after: number, before: number) => goodRows.filter((row) => row.row > after && row.row < before);
  const pair = (indexes: number[], among: SlidesheetSurveyRow[]) => {
    among.forEach((row, at) => {
      rowOf[indexes[at]!] = row;
      pairs.push({ index: indexes[at]!, row });
    });
  };
  const first = anchors[0]!;
  const last = anchors[anchors.length - 1]!;

  const leadingRows = goodBetween(0, first.row.row);
  const leadingStations = Array.from({ length: first.index }, (_, index) => index);
  if (leadingRows.length > 0 && leadingStations.length > 0) {
    if (leadingStations.length < leadingRows.length) return reject(list[leadingStations[0]!]!.md, `has fewer surveys above row ${first.row.row} than the Slidesheet has rows.`);
    pair(leadingStations.slice(-leadingRows.length), leadingRows);
  }
  ignored += leadingStations.filter((index) => rowOf[index] == null).length;

  for (let at = 0; at + 1 < anchors.length; at += 1) {
    const from = anchors[at]!;
    const to = anchors[at + 1]!;
    const between = goodBetween(from.row.row, to.row.row);
    const inside = Array.from({ length: to.index - from.index - 1 }, (_, offset) => from.index + 1 + offset);
    if (inside.length === 0) continue;
    if (between.length !== inside.length) {
      const detail = between.length === 0
        ? "has no Slidesheet row to take it."
        : `has ${inside.length} survey(s) between MD ${fixed2(list[from.index]!.md)} and ${fixed2(list[to.index]!.md)} where the Slidesheet has ${between.length}.`;
      return reject(list[inside[0]!]!.md, detail);
    }
    pair(inside, between);
  }

  const trailingRows = goodRows.filter((row) => row.row > last.row.row);
  const afterLast = Array.from({ length: list.length - last.index - 1 }, (_, offset) => last.index + 1 + offset);
  let tail = afterLast;
  if (trailingRows.length > 0 && afterLast.length > 0) {
    if (afterLast.length < trailingRows.length) {
      return reject(list[afterLast[0]!]!.md, `has fewer surveys below row ${last.row.row} than the Slidesheet has good rows.`);
    }
    pair(afterLast.slice(0, trailingRows.length), trailingRows);
    tail = afterLast.slice(trailingRows.length);
  }

  const matched: MatchedSurvey[] = [];
  for (const { index, row } of anchors) matched.push({ station: list[index]!, target: row });
  for (const { index, row } of pairs) {
    const depth = depthForStation(list[index]!, coordinate, ownOffset(row) ?? sheetOffset);
    if (depth == null) return reject(list[index]!.md, "the Slidesheet survey offset could not be determined.");
    matched.push({ station: list[index]!, target: row, depth });
  }

  // 3. The newer surveys go on the rows after the last good one, in order.
  const tailRows: Array<{ station: SurveyStation; row: number; had: SlidesheetSurveyRow | undefined; depth: number | null }> = [];
  if (lastGoodRow + tail.length > SS_LAST) {
    const rejected = list[tail[SS_LAST - lastGoodRow]!];
    return reject(rejected?.md ?? list[tail[0]!]!.md, `would exceed Slidesheet row ${SS_LAST}.`);
  }
  for (let at = 0; at < tail.length; at += 1) {
    const station = list[tail[at]!]!;
    const row = lastGoodRow + 1 + at;
    const had = byRow.get(row);
    const current = had ? depthAt(had, coordinate) : null;
    const depth = depthForStation(station, coordinate, ownOffset(had) ?? sheetOffset);
    const bitAgrees = coordinate !== "survey-depth" || station.bit == null || (had?.d != null && closeSurveyDepth(had.d, station.bit));
    const agrees = current != null && closeSurveyDepth(current, station.md) && bitAgrees;
    if (depth == null) return reject(station.md, "the Slidesheet survey offset could not be determined.");
    const dCell = numeric(book.value(SS, row, COL_D));
    if (!agrees && book.hasFormula(SS, row, COL_D)) return reject(station.md, `cannot use Slidesheet row ${row}.`);
    if (!agrees && current == null && dCell != null) return reject(station.md, `cannot use Slidesheet row ${row}.`);
    tailRows.push({ station, row, had, depth: agrees ? null : depth });
  }

  const tvdMismatches = tvdCorrectionMismatches(book, matched, lastGoodRow);
  if (tvdMismatches == null) {
    const correction = matched.find(({ station, target }) =>
      target.row <= lastGoodRow &&
      (numeric(book.value(SS, target.row, COL_F)) !== station.inc || numeric(book.value(SS, target.row, COL_G)) !== station.azm));
    return reject(correction?.station.md ?? null, "a past-survey correction has no comparable FieldCap and Slidesheet TVD.");
  }
  if (tvdMismatches.length > 0 && !options.allowTvdMismatch) {
    return {
      ...emptySurveySyncResult(),
      coordinate,
      matched: matched.length,
      ignored,
      skippedPlaceholders: skipped,
      needsConfirmation: true,
      tvdMismatches,
    };
  }

  let changed = 0;
  let unchanged = 0;
  let appended = 0;
  const depthCorrections: SurveyDepthCorrection[] = [];
  const writeSurvey = (station: SurveyStation, row: number, depth: number | null, had: SlidesheetSurveyRow | undefined): boolean => {
    let touched = false;
    if (depth != null) {
      book.setValue(SS, row, COL_D, depth);
      touched = true;
      if (had?.d != null) depthCorrections.push({ row, md: station.md, from: had.d, to: depth });
    }
    if (numeric(book.value(SS, row, COL_F)) !== station.inc) {
      book.setValue(SS, row, COL_F, station.inc);
      touched = true;
    }
    if (numeric(book.value(SS, row, COL_G)) !== station.azm) {
      book.setValue(SS, row, COL_G, station.azm);
      touched = true;
    }
    if (book.fill(SS, row, COL_F) !== GOOD_SURVEY_FILL || book.fill(SS, row, COL_G) !== GOOD_SURVEY_FILL) {
      book.setFill(SS, { row, col: COL_F, row2: row, col2: COL_G }, GOOD_SURVEY_FILL);
      touched = true;
    }
    return touched;
  };

  for (const { station, target, depth } of matched) {
    if (writeSurvey(station, target.row, depth ?? null, target)) changed += 1;
    else unchanged += 1;
  }
  for (const { station, row, had, depth } of tailRows) {
    if (had == null || depthAt(had, coordinate) == null) appended += 1;
    if (writeSurvey(station, row, depth, had)) changed += 1;
    else unchanged += 1;
  }
  const offsetChanges = coordinate === "survey-depth"
    ? followFieldCapOffsets(book, [
      ...matched.map(({ station, target }) => ({ row: target.row, station })),
      ...tailRows.map(({ station, row }) => ({ row, station })),
    ], depthCorrections)
    : [];
  depthCorrections.sort((left, right) => left.row - right.row);

  const newestWritten = tailRows.length > 0 ? tailRows[tailRows.length - 1]!.row : 0;
  const newest = newestWritten > 0 ? newestWritten : lastSurveyRow(book);
  const staleMarker = rows.some((row) => row.row !== newest && book.fill(SS, row.row, 10) === GOOD_SURVEY_FILL);
  if (changed > 0 || offsetChanges.length > 0 || staleMarker) highlightNewestSurvey(book, newest > 0 ? newest : undefined);

  return {
    ...emptySurveySyncResult(),
    coordinate,
    matched: matched.length + tailRows.filter((entry) => entry.had != null && depthAt(entry.had, coordinate) != null).length,
    changed,
    unchanged,
    appended,
    ignored,
    skippedPlaceholders: skipped,
    tvdMismatches,
    depthCorrections,
    offsetChanges,
  };
}
