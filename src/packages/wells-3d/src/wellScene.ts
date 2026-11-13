import { minCurveStep } from "../../dd-math/src/angles";
import { checkAcReport, interpPlan, type AcPlanPoint } from "../../workbook-commands/src/acCheck";
import { readAcDetail, wellsMatch, type AcDetailStation } from "../../workbook-commands/src/acDetail";
import { acConcernCapacity, concernWells, readAcHits, type AcHit } from "../../workbook-commands/src/acImport";
import { isSummaryRow } from "../../workbook-commands/src/planGauge";
import { SURVEY_SHEET } from "../../workbook-commands/src/planImport";
import { numeric, type SheetBook } from "../../workbook-commands/src/sheetBook";
import { isGoodSurveyRow, SS_SHEET } from "../../workbook-commands/src/slidesheet";

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
  marker: { x: number; y: number; z: number } | null;
  join: { x: number; y: number; z: number } | null;
}

export interface WellScene {
  plan: PathArrays;
  drilled: PathArrays | null;
  offsets: OffsetPath[];
  /** Concern wells the report gave no path for. */
  missing: string[];
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
  const last = book.lastUsedRow(SURVEY_SHEET, 1);
  const out: AcPlanPoint[] = [];
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

function readDrilled(book: SheetBook): Sample[] {
  if (!book.hasSheet(SS_SHEET)) return [];
  let prev = { md: 0, inc: 0, azi: 0 };
  let n = 0;
  let e = 0;
  let tvd = 0;
  const out: Sample[] = [];
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

function hashScene(plan: PathArrays, drilled: PathArrays | null, offsets: OffsetPath[], missing: string[], badge: string): string {
  let hash = hashText(badge);
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
    hash ^= hashText(`${offset.name}:${offset.concern ? "Y" : "N"}`);
    mix(offset);
  }
  hash ^= hashText(missing.join("|"));
  return (hash >>> 0).toString(16).padStart(8, "0");
}

function closestStation(stations: AcDetailStation[], md: number): AcDetailStation | null {
  let best: AcDetailStation | null = null;
  let gap = Infinity;
  for (const station of stations) {
    const next = Math.abs(station.refMd - md);
    if (next < gap) {
      gap = next;
      best = station;
    }
  }
  return best;
}

function offsetPath(name: string, stations: AcDetailStation[], hit: AcHit | null, concern: boolean, origin: AcPlanPoint, plan: AcPlanPoint[]): OffsetPath {
  const ordered = [...stations].sort((a, b) => a.refMd - b.refMd);
  const bandSf = hit?.sf ?? ordered.reduce((lowest, station) => (station.sf > 0 && station.sf < lowest ? station.sf : lowest), Infinity);
  const samples: Sample[] = ordered.map((station) => ({
    md: station.refMd,
    n: station.n,
    e: station.e,
    tvd: station.offTvd,
    c2c: station.c2c,
    sf: station.sf,
  }));
  const markAt = closestStation(ordered, hit?.refMd ?? ordered.reduce((best, station) => (station.sf > 0 && (best == null || station.sf < best.sf) ? station : best), null as AcDetailStation | null)?.refMd ?? ordered[0]!.refMd);
  let marker: OffsetPath["marker"] = null;
  let join: OffsetPath["join"] = null;
  if (markAt) {
    marker = localOf({ md: markAt.refMd, n: markAt.n, e: markAt.e, tvd: markAt.offTvd, c2c: markAt.c2c, sf: markAt.sf }, origin);
    const onPlan = interpPlan(plan, hit?.refMd ?? markAt.refMd);
    if (onPlan) join = localOf({ md: onPlan.md, n: onPlan.n, e: onPlan.e, tvd: onPlan.tvd, c2c: null, sf: null }, origin);
  }
  return {
    ...pack(samples, origin),
    name,
    concern,
    bandSf: Number.isFinite(bandSf) ? bandSf : 99,
    color: bandColor(Number.isFinite(bandSf) ? bandSf : 99),
    marker,
    join,
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
  const byWell = new Map<string, AcDetailStation[]>();
  for (const station of stations) {
    const key = [...byWell.keys()].find((name) => wellsMatch(name, station.well));
    const name = key ?? station.well;
    const list = byWell.get(name);
    if (list) list.push(station);
    else byWell.set(name, [station]);
  }
  const offsets = [...byWell.entries()].map(([name, rows]) => {
    const hit = hits.find((item) => wellsMatch(item.well, name)) ?? null;
    const concern = shown.some((item) => wellsMatch(item.well, name));
    return offsetPath(name, rows, hit, concern, origin, plan);
  });
  const missing = shown.filter((hit) => !offsets.some((offset) => wellsMatch(offset.name, hit.well))).map((hit) => hit.well);
  const drilledSamples = readDrilled(book);
  const drilled = drilledSamples.length >= 2 ? pack(drilledSamples, origin) : null;
  const check = checkAcReport(shown, stations, plan);
  const badge = shown.length === 0
    ? (stations.length === 0 ? "No offset paths in the report" : "No concern wells")
    : check.badge;
  const badgeOk = shown.length === 0 ? true : check.ok;
  const planPath = pack(plan.map((point) => ({ ...point, c2c: null, sf: null })), origin);
  return {
    plan: planPath,
    drilled,
    offsets,
    missing,
    badge,
    badgeOk,
    hash: hashScene(planPath, drilled, offsets, missing, badge),
  };
}
