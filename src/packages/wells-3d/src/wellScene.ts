import { minCurveStep } from "../../dd-math/src/angles";
import { integrateLocalNE } from "../../dd-math/src/plan";
import { checkAcReport, interpPlan, stationInFrame, type AcPlanPoint } from "../../workbook-commands/src/acCheck";
import { readAcDetail, wellsMatch, type AcDetailStation } from "../../workbook-commands/src/acDetail";
import { compareHeaders, readDatum } from "../../workbook-commands/src/acHeader";
import { acConcernCapacity, concernWells, readAcHits, type AcHit } from "../../workbook-commands/src/acImport";
import { isSummaryRow } from "../../workbook-commands/src/planGauge";
import { SURVEY_SHEET } from "../../workbook-commands/src/planImport";
import { cellText, numeric, type SheetBook } from "../../workbook-commands/src/sheetBook";
import { isGoodSurveyRow, loadNamedTargets, SS_SHEET } from "../../workbook-commands/src/slidesheet";
import { loadFormations } from "../../corridor-3d/src/corridorScene";
import { ringRadii } from "./ringLabel";

// East, North, and up, in metres from the first plan station. Up is the opposite of TVD.
// 1:1, so a metre of hole is a metre on every axis.

export interface PathArrays {
  /** Packed x, y, z. */
  xyz: Float32Array;
  md: Float32Array;
  tvd: Float32Array;
  /** Printed centre-to-centre. NaN where the path has none. */
  c2c: Float32Array;
  /** Printed separation factor. NaN where the path has none. */
  sf: Float32Array;
}

export interface OffsetPath extends PathArrays {
  name: string;
  concern: boolean;
  /** Separation factor the line is coloured by. */
  bandSf: number;
  color: string;
  /** The Summary site heading. Empty when the report gave none. */
  site: string;
  /** The well is listed under the report's reference site, the active well's pad. */
  onPad: boolean;
  /** The lowest printed centre-to-centre. 0 when the report gave none. */
  minC2c: number;
  /** The offset's surface, from the active well. Null when the report has no row at the surface. */
  wellhead: Wellhead | null;
  /**
   * Stations that start a new run. The report only gives the offset's closest point for each
   * reference station. A normal survey step, and a longer step that is still one turn of the
   * well, is drawn. A step whose straight line is much shorter than the measured depth is left
   * open: that stretch was never surveyed, and a chord would cut the corner. Sorted, never 0.
   */
  breaks: number[];
}

/** Metres north and east of the active well's slot, and the compass bearing of that point. */
export interface Wellhead {
  north: number;
  east: number;
  distance: number;
  /** Degrees clockwise from north, 0 to 359. */
  bearing: number;
}

/** The report header the view shows, and the verdict of the check against the plan. */
export interface WellReport {
  stored: boolean;
  referenceWell: string;
  design: string;
  referenceSite: string;
  coordinatesRelativeTo: string;
  north: string;
  tvdReference: string;
  offsetTvdReference: string;
  reportDate: string;
  verdict: string;
  verdictOk: boolean;
}

/** A formation top from the plan, in the same metres as the well. */
export interface WellFormation {
  name: string;
  lith: string;
  /** True vertical depth on the plan datum, metres. */
  tvd: number;
  /** Scene up, metres. Up is the opposite of TVD, from the first plan station. */
  z: number;
}

/** A named plan target. The rings sit on the plan, square to the wellbore. */
export interface WellTarget {
  name: string;
  md: number;
  tvd: number;
  x: number;
  y: number;
  z: number;
  /** Outer ring, 20 m. The inner ring is 10 m. */
  radius: number;
  outer: Float32Array;
  inner: Float32Array;
}

export interface WellScene {
  plan: PathArrays;
  drilled: PathArrays | null;
  offsets: OffsetPath[];
  targets: WellTarget[];
  /** Formation tops from the plan. The walls label the ones that fall on the box. */
  formations: WellFormation[];
  /** Concern wells the report gave no path for. */
  missing: string[];
  /** Report rows not drawn because they do not sit where their printed C2C says. */
  leftOut: number;
  report: WellReport;
  badge: string;
  badgeOk: boolean;
  hash: string;
}

const SURVEY_FIRST = 13;
const SURVEY_LAST = 520;

/** The ink already used for a separation factor: red below 1.5, yellow below 2, green otherwise. */
export function bandColor(sf: number): string {
  if (sf < 1.5) return "#C62828";
  if (sf < 2) return "#F9A825";
  return "#2E7D32";
}

interface Local {
  x: number;
  y: number;
  z: number;
}

interface Sample {
  md: number;
  n: number;
  e: number;
  tvd: number;
  c2c: number | null;
  sf: number | null;
}

function localOf(sample: Sample, origin: AcPlanPoint): Local {
  return {
    x: sample.e - origin.e,
    y: sample.n - origin.n,
    z: -(sample.tvd - origin.tvd),
  };
}

function pack(samples: Sample[], origin: AcPlanPoint): PathArrays {
  const xyz = new Float32Array(samples.length * 3);
  const md = new Float32Array(samples.length);
  const tvd = new Float32Array(samples.length);
  const c2c = new Float32Array(samples.length);
  const sf = new Float32Array(samples.length);
  samples.forEach((sample, i) => {
    const at = localOf(sample, origin);
    xyz[i * 3] = at.x;
    xyz[i * 3 + 1] = at.y;
    xyz[i * 3 + 2] = at.z;
    md[i] = sample.md;
    tvd[i] = sample.tvd;
    c2c[i] = sample.c2c ?? Number.NaN;
    sf[i] = sample.sf ?? Number.NaN;
  });
  return { xyz, md, tvd, c2c, sf };
}

function readPlan(book: SheetBook): AcPlanPoint[] {
  if (!book.hasSheet(SURVEY_SHEET)) return [];
  const cols: Partial<Record<"md" | "inc" | "azi" | "tvd" | "ns" | "ew", number>> = {};
  for (let c = 1; c <= 20; c += 1) {
    const head = cellText(book.value(SURVEY_SHEET, 2, c)).toUpperCase();
    if (head === "MD") cols.md = c;
    else if (head === "INC") cols.inc = c;
    else if (head === "AZI" || head === "AZM") cols.azi = c;
    else if (head === "TVD") cols.tvd = c;
    else if (head === "NS") cols.ns = c;
    else if (head === "EW") cols.ew = c;
  }
  const colMd = cols.md ?? 1;
  const colInc = cols.inc ?? 2;
  const colAzi = cols.azi ?? 3;
  const colTvd = cols.tvd ?? 4;
  const colNs = cols.ns ?? 5;
  const colEw = cols.ew ?? 6;
  const last = book.lastUsedRow(SURVEY_SHEET, colMd);
  const rows: Array<AcPlanPoint & { inc: number | null; azi: number | null; missingNE: boolean }> = [];
  for (let row = 3; row <= last; row += 1) {
    const md = numeric(book.value(SURVEY_SHEET, row, colMd));
    const tvd = numeric(book.value(SURVEY_SHEET, row, colTvd));
    const inc = numeric(book.value(SURVEY_SHEET, row, colInc));
    const azi = numeric(book.value(SURVEY_SHEET, row, colAzi));
    const n = numeric(book.value(SURVEY_SHEET, row, colNs));
    const e = numeric(book.value(SURVEY_SHEET, row, colEw));
    if (md == null || tvd == null) continue;
    if ((n == null || e == null) && (inc == null || azi == null)) continue;
    if (rows.length > 0 && !(md > rows[rows.length - 1]!.md)) continue;
    rows.push({ md, n: n ?? 0, e: e ?? 0, tvd, inc, azi, missingNE: n == null || e == null });
  }
  if (rows.length > 0 && rows.every((row) => row.inc != null && row.azi != null)) {
    const startN = rows[0]!.md > 0.01 && !rows[0]!.missingNE ? rows[0]!.n : 0;
    const startE = rows[0]!.md > 0.01 && !rows[0]!.missingNE ? rows[0]!.e : 0;
    const ne = integrateLocalNE(
      rows.map((row) => row.md),
      rows.map((row) => row.inc!),
      rows.map((row) => row.azi!),
      startN,
      startE,
    );
    let storedOk = !rows.some((row) => row.missingNE);
    if (storedOk && rows[0]!.md <= 0.01 && (Math.abs(rows[0]!.n) > 0.1 || Math.abs(rows[0]!.e) > 0.1)) {
      storedOk = false;
    } else if (storedOk) {
      let maxErr = 0;
      for (let i = 0; i < rows.length; i += 1) {
        const err = Math.hypot(rows[i]!.n - ne.ns[i]!, rows[i]!.e - ne.ew[i]!);
        if (err > maxErr) maxErr = err;
      }
      if (maxErr > 2) storedOk = false;
    }
    if (!storedOk) {
      for (let i = 0; i < rows.length; i += 1) {
        rows[i]!.n = ne.ns[i]!;
        rows[i]!.e = ne.ew[i]!;
      }
    }
  }
  return rows.map(({ md, n, e, tvd }) => ({ md, n, e, tvd }));
}

function cellNumOr(book: SheetBook, row: number, col: number, fallback = 0): number {
  const value = book.value(SS_SHEET, row, col);
  const parsed = value == null || value === "" ? null : numeric(value);
  return parsed == null ? fallback : parsed;
}

function readDrilled(book: SheetBook): Sample[] {
  if (!book.hasSheet(SS_SHEET)) return [];
  const tieRow = SURVEY_FIRST - 1;
  const tieMd = cellNumOr(book, tieRow, 5, 0);
  let prev = {
    md: tieMd,
    inc: cellNumOr(book, tieRow, 6, 0),
    azi: cellNumOr(book, tieRow, 7, 0),
  };
  let n = cellNumOr(book, tieRow, 38, 0);
  let e = cellNumOr(book, tieRow, 39, 0);
  let tvd = cellNumOr(book, tieRow, 8, tieMd);
  const out: Sample[] = [];
  if (tieMd > 0) out.push({ md: tieMd, n, e, tvd, c2c: null, sf: null });
  for (let row = SURVEY_FIRST; row <= SURVEY_LAST; row += 1) {
    if (isSummaryRow(book, row) || !isGoodSurveyRow(book, row)) continue;
    const md = numeric(book.value(SS_SHEET, row, 5));
    const inc = numeric(book.value(SS_SHEET, row, 6));
    const azi = numeric(book.value(SS_SHEET, row, 7));
    if (md == null || inc == null || azi == null || !(md > prev.md)) continue;
    const step = minCurveStep(prev.md, prev.inc, prev.azi, md, inc, azi);
    n += step.n;
    e += step.e;
    tvd += step.tvd;
    prev = { md, inc, azi };
    out.push({ md, n, e, tvd, c2c: null, sf: null });
  }
  return out;
}

function hashText(text: string): number {
  let hash = 2_166_136_261;
  for (let i = 0; i < text.length; i += 1) {
    hash ^= text.charCodeAt(i);
    hash = Math.imul(hash, 1_677_7619);
  }
  return hash >>> 0;
}

function hashScene(plan: PathArrays, drilled: PathArrays | null, offsets: OffsetPath[], targets: WellTarget[], formations: WellFormation[], missing: string[], report: WellReport, badge: string): string {
  let hash = hashText(`${badge}|${report.verdict}|${report.referenceWell}|${report.design}|${report.referenceSite}`);
  const mix = (path: PathArrays | null) => {
    if (!path) return;
    for (let i = 0; i < path.xyz.length; i += 1) {
      hash ^= Math.round(path.xyz[i]! * 1000);
      hash = Math.imul(hash, 1_677_7619);
    }
  };
  mix(plan);
  mix(drilled);
  for (const offset of offsets) {
    hash ^= hashText(`${offset.name}:${offset.concern ? "Y" : "N"}:${offset.site}:${offset.minC2c}:${offset.wellhead?.bearing ?? ""}:${offset.breaks.join(",")}`);
    mix(offset);
  }
  for (const target of targets) {
    hash ^= hashText(`${target.name}:${target.md}`);
    hash = Math.imul(hash, 1_677_7619);
  }
  for (const formation of formations) {
    hash ^= hashText(`${formation.name}:${formation.tvd}:${formation.lith}`);
    hash = Math.imul(hash, 1_677_7619);
  }
  hash ^= hashText(missing.join("|"));
  return (hash >>> 0).toString(16).padStart(8, "0");
}

/** A normal survey interval. A step this short is one leg, even when the coordinates wander a little. */
const KNOWN_STEP_M = 90;
/**
 * Chord divided by the measured step. 0.85 is about a 110° turn, which is still one leg
 * between two known stations. A smaller ratio is a jump across hole the report never gave.
 */
const CURVE_RATIO = 0.85;
function chordOf(a: AcDetailStation, b: AcDetailStation): number {
  return Math.hypot(b.n - a.n, b.e - a.e, b.offTvd - a.offTvd);
}

function sfRank(station: AcDetailStation): number {
  return station.sf > 0 ? station.sf : Number.POSITIVE_INFINITY;
}

/**
 * The offset's own points in the offset's own depth order. The report lists each reference station
 * with the offset's closest point, and that point runs back and forth along the offset as the
 * reference moves. Ordered by the reference, the line doubled back on itself. Where several
 * reference stations share one offset point, the tightest separation factor comes last, which is
 * the one a hover lands on.
 */
function alongOffset(stations: AcDetailStation[]): AcDetailStation[] {
  return [...stations].sort((a, b) => {
    if (a.offMd !== b.offMd) return a.offMd - b.offMd;
    const left = sfRank(a);
    const right = sfRank(b);
    if (left === right) return a.refMd - b.refMd;
    return left < right ? 1 : -1;
  });
}

/** Where the report gives no path for a stretch of the offset, and the stretch is not one turn of the well. */
function runBreaks(stations: AcDetailStation[]): number[] {
  const out: number[] = [];
  for (let i = 1; i < stations.length; i += 1) {
    const previous = stations[i - 1]!;
    const next = stations[i]!;
    const along = next.offMd - previous.offMd;
    const chord = chordOf(previous, next);
    // Two reference rows can share a depth and still sit in different places. That spur is not a leg.
    if (along < 1 && chord > 5) {
      out.push(i);
      continue;
    }
    if (along <= KNOWN_STEP_M) continue;
    if (chord >= along * CURVE_RATIO - 1) continue;
    out.push(i);
  }
  return out;
}

/** A row this close to the offset's surface is the wellhead. Deeper, and the report never gave the surface. */
const SURFACE_M = 1;

function squash(text: string): string {
  return text.toLowerCase().replace(/\s+/g, " ").trim();
}

function siteName(stations: AcDetailStation[]): string {
  for (const station of stations) {
    const site = station.site.trim();
    if (site !== "") return site;
  }
  return "";
}

function lowestC2c(stations: AcDetailStation[]): number {
  let lowest = Infinity;
  for (const station of stations) if (station.c2c > 0 && station.c2c < lowest) lowest = station.c2c;
  return Number.isFinite(lowest) ? lowest : 0;
}

function wellheadOf(stations: AcDetailStation[]): Wellhead | null {
  const surface = stations.reduce<AcDetailStation | null>((best, station) => (!best || station.offMd < best.offMd ? station : best), null);
  if (!surface || surface.offMd > SURFACE_M) return null;
  const distance = surface.c2c > 0 ? surface.c2c : Math.hypot(surface.n, surface.e, surface.offTvd);
  const degrees = Math.atan2(surface.e, surface.n) * 180 / Math.PI;
  return {
    north: surface.n,
    east: surface.e,
    distance,
    bearing: Math.round((degrees % 360 + 360) % 360) % 360,
  };
}

function offsetPath(name: string, stations: AcDetailStation[], hit: AcHit | null, concern: boolean, origin: AcPlanPoint, referenceSite: string): OffsetPath {
  const ordered = alongOffset(stations);
  const bandSf = hit?.sf ?? ordered.reduce((lowest, station) => (station.sf > 0 && station.sf < lowest ? station.sf : lowest), Infinity);
  const site = siteName(ordered);
  const samples: Sample[] = ordered.map((station) => ({
    md: station.refMd,
    n: station.n,
    e: station.e,
    tvd: station.offTvd,
    c2c: station.c2c,
    sf: station.sf > 0 ? station.sf : null,
  }));
  return {
    ...pack(samples, origin),
    name,
    concern,
    bandSf: Number.isFinite(bandSf) ? bandSf : 99,
    color: bandColor(Number.isFinite(bandSf) ? bandSf : 99),
    site,
    onPad: site !== "" && squash(site) === squash(referenceSite),
    minC2c: lowestC2c(ordered),
    wellhead: wellheadOf(ordered),
    breaks: runBreaks(ordered),
  };
}

function blankReport(verdict: string): WellReport {
  return {
    stored: false,
    referenceWell: "",
    design: "",
    referenceSite: "",
    coordinatesRelativeTo: "",
    north: "",
    tvdReference: "",
    offsetTvdReference: "",
    reportDate: "",
    verdict,
    verdictOk: true,
  };
}

/** The stored header, and the sentence that replaces the frame badge when the report is not the plan's. */
function reportOf(book: SheetBook): { report: WellReport; replaceBadge: string } {
  const datum = readDatum(book);
  if (!datum.reportStored) return { report: blankReport("Run Import AC again to store the report header"), replaceBadge: "" };
  const check = compareHeaders(datum.report, datum.planStored ? datum.plan : null);
  return {
    report: {
      stored: true,
      referenceWell: datum.report.referenceWell,
      design: datum.report.referenceDesign,
      referenceSite: datum.report.referenceSite,
      coordinatesRelativeTo: datum.report.coordinatesRelativeTo || datum.report.localCoordinateReference,
      north: datum.report.northReference,
      tvdReference: datum.report.tvdReference,
      offsetTvdReference: datum.report.offsetTvdReference,
      reportDate: datum.report.reportDate,
      verdict: check.verdict,
      verdictOk: check.ok,
    },
    replaceBadge: check.ok ? "" : check.verdict,
  };
}

/** The whole plan, the drilled path, and every offset well the report gave a path for. Null until the plan has two stations. */
export function buildWellScene(book: SheetBook): WellScene | null {
  const plan = readPlan(book);
  if (plan.length < 2) return null;
  const origin = plan[0]!;
  const hits = readAcHits(book);
  const shown = concernWells(hits, acConcernCapacity(book));
  const stations = readAcDetail(book);
  // A row is drawn only when its printed C2C matches the plan. A row that does not match is a row
  // that was read wrong, and drawn it is a straight line to nowhere.
  const proven = stations.filter((station) => stationInFrame(plan, station));
  const leftOut = stations.length - proven.length;
  const byWell = new Map<string, AcDetailStation[]>();
  for (const station of proven) {
    const key = [...byWell.keys()].find((name) => wellsMatch(name, station.well));
    const name = key ?? station.well;
    const list = byWell.get(name);
    if (list) list.push(station);
    else byWell.set(name, [station]);
  }
  const { report, replaceBadge } = reportOf(book);
  const offsets = [...byWell.entries()].map(([name, rows]) => {
    const hit = hits.find((item) => wellsMatch(item.well, name)) ?? null;
    const concern = shown.some((item) => wellsMatch(item.well, name));
    return offsetPath(name, rows, hit, concern, origin, report.referenceSite);
  });
  const missing = shown.filter((hit) => !offsets.some((offset) => wellsMatch(offset.name, hit.well))).map((hit) => hit.well);
  const drilledSamples = readDrilled(book);
  const drilled = drilledSamples.length >= 2 ? pack(drilledSamples, origin) : null;
  const check = checkAcReport(shown, stations, plan);
  const leftOutNote = leftOut > 0 ? `, ${leftOut} row${leftOut === 1 ? "" : "s"} left out` : "";
  const frameBadge = shown.length === 0
    ? (stations.length === 0 ? "No offset paths in the report" : "No concern wells")
    : `${check.badge}${leftOutNote}`;
  const badge = replaceBadge !== "" ? replaceBadge : frameBadge;
  const badgeOk = replaceBadge !== "" ? false : shown.length === 0 ? true : check.ok;
  const planPath = pack(plan.map((point) => ({ ...point, c2c: null, sf: null })), origin);
  const targets = planTargets(book, plan, origin);
  const formations = formationsOf(book, origin);
  return {
    plan: planPath,
    drilled,
    offsets,
    targets,
    formations,
    missing,
    leftOut,
    report,
    badge,
    badgeOk,
    hash: hashScene(planPath, drilled, offsets, targets, formations, missing, report, badge),
  };
}

/** _OC_Formations, the same tops the email picture writes on its shaft walls. */
function formationsOf(book: SheetBook, origin: AcPlanPoint): WellFormation[] {
  return loadFormations(book)
    .filter((fm) => fm.name.trim() !== "" && Number.isFinite(fm.tvd))
    .map((fm) => ({
      name: fm.name,
      lith: fm.lith,
      tvd: fm.tvd,
      z: -(fm.tvd - origin.tvd),
    }));
}

const WINDOW_FIRST = 2;
const WINDOW_LAST = 5;
const WINDOW_MD = 21;
const WINDOW_NAME = 25;
const RING_POINTS = 32;

function namedOnPlan(book: SheetBook): Array<{ name: string; md: number }> {
  const named = loadNamedTargets(book);
  if (named.length > 0) return named.map((target) => ({ name: target.name, md: target.md }));
  if (!book.hasSheet(SS_SHEET)) return [];
  const out: Array<{ name: string; md: number }> = [];
  for (let row = WINDOW_FIRST; row <= WINDOW_LAST; row += 1) {
    const name = cellText(book.value(SS_SHEET, row, WINDOW_NAME));
    const md = numeric(book.value(SS_SHEET, row, WINDOW_MD));
    if (name === "" || md == null) continue;
    out.push({ name, md });
  }
  return out;
}

function unit(vector: Local): Local {
  const length = Math.hypot(vector.x, vector.y, vector.z);
  if (length < 1e-6) return { x: 0, y: 0, z: 1 };
  return { x: vector.x / length, y: vector.y / length, z: vector.z / length };
}

function cross(a: Local, b: Local): Local {
  return {
    x: a.y * b.z - a.z * b.y,
    y: a.z * b.x - a.x * b.z,
    z: a.x * b.y - a.y * b.x,
  };
}

/** A circle in the plane square to the wellbore, closed by repeating the first point. */
function ringPoints(center: Local, tangent: Local, radius: number): Float32Array {
  const along = unit(tangent);
  const helper = Math.abs(along.z) < 0.95 ? { x: 0, y: 0, z: 1 } : { x: 1, y: 0, z: 0 };
  const across = unit(cross(along, helper));
  const around = cross(along, across);
  const out = new Float32Array((RING_POINTS + 1) * 3);
  for (let i = 0; i <= RING_POINTS; i += 1) {
    const angle = (2 * Math.PI * i) / RING_POINTS;
    const cos = Math.cos(angle);
    const sin = Math.sin(angle);
    out[i * 3] = center.x + radius * (cos * across.x + sin * around.x);
    out[i * 3 + 1] = center.y + radius * (cos * across.y + sin * around.y);
    out[i * 3 + 2] = center.z + radius * (cos * across.z + sin * around.z);
  }
  return out;
}

function planTargets(book: SheetBook, plan: AcPlanPoint[], origin: AcPlanPoint): WellTarget[] {
  const { outer: radius, inner } = ringRadii();
  const out: WellTarget[] = [];
  for (const target of namedOnPlan(book)) {
    const at = interpPlan(plan, target.md);
    if (!at) continue;
    const before = interpPlan(plan, target.md - 6) ?? plan[0]!;
    const after = interpPlan(plan, target.md + 6) ?? plan[plan.length - 1]!;
    const center = localOf({ ...at, c2c: null, sf: null }, origin);
    const tangent = {
      x: localOf({ ...after, c2c: null, sf: null }, origin).x - localOf({ ...before, c2c: null, sf: null }, origin).x,
      y: localOf({ ...after, c2c: null, sf: null }, origin).y - localOf({ ...before, c2c: null, sf: null }, origin).y,
      z: localOf({ ...after, c2c: null, sf: null }, origin).z - localOf({ ...before, c2c: null, sf: null }, origin).z,
    };
    out.push({
      name: target.name,
      md: at.md,
      tvd: at.tvd,
      x: center.x,
      y: center.y,
      z: center.z,
      radius,
      outer: ringPoints(center, tangent, radius),
      inner: ringPoints(center, tangent, inner),
    });
  }
  return out;
}
