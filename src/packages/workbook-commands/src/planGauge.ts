import { minCurveStep } from "../../dd-math/src/angles";
import { integrateLocalNE } from "../../dd-math/src/plan";
import { parseToolface, walkUnseenToBit, type SlideInterval } from "../../dd-math/src/project";
import { cellText, fixed2, isNumber, numeric, SheetBook } from "./sheetBook";
import { isGoodSurveyRow, SS_SHEET, vbaVal } from "./slidesheet";

// Port of MDL_PlanGauge.RenderPlanGaugeCore (the Slidesheet AA1:AC7 proximity dial) and
// MDL_WellborePos.UpdateWellborePosition (Data "Position of Wellbore").

const PLAN_SHEET = "_OC_Survey";
const SURV_FIRST = 13;
const SURV_LAST = 520;
const CROSSOVER_INC = 5;
const WP_FIRST = 14;
const WP_LAST = 33;
const WP_COL_MD = 29;
const WP_COL_TVD = 30;
const INC_LATERAL = 80;

export type GaugeMode = "GRAV" | "MAG";

export interface GaugeModel {
  hasData: boolean;
  caption: string;
  mode: GaugeMode;
  ax: number;
  ay: number;
  showPtb: boolean;
  bx: number;
  by: number;
  hasBand: boolean;
  yTop: number;
  yBot: number;
  scale: number;
}

export interface Plan {
  md: number[];
  inc: number[];
  azi: number[];
  tvd: number[];
  ns: number[];
  ew: number[];
}

export interface PlanPoint {
  n: number;
  e: number;
  v: number;
  azi: number;
  inc: number;
}

const rad = (d: number) => d * Math.PI / 180;

function clamp(v: number, lo: number, hi: number): number {
  return v < lo ? lo : v > hi ? hi : v;
}

export function loadPlan(book: SheetBook): Plan | null {
  if (!book.hasSheet(PLAN_SHEET)) return null;
  const cols: Record<string, number> = {};
  for (let c = 1; c <= 20; c += 1) {
    const head = cellText(book.value(PLAN_SHEET, 2, c)).toUpperCase();
    if (head === "MD") cols.md = c;
    else if (head === "INC") cols.inc = c;
    else if (head === "AZI" || head === "AZM") cols.azi = c;
    else if (head === "TVD") cols.tvd = c;
    else if (head === "NS") cols.ns = c;
    else if (head === "EW") cols.ew = c;
  }
  const { md, inc, azi, tvd, ns, ew } = cols;
  if (!md || !inc || !azi || !tvd || !ns || !ew) return null;
  const plan: Plan = { md: [], inc: [], azi: [], tvd: [], ns: [], ew: [] };
  let lastR = 0;
  for (let r = 3; r <= 5000; r += 1) if (cellText(book.value(PLAN_SHEET, r, md)) !== "") lastR = r;
  for (let r = 3; r <= lastR; r += 1) {
    const raw = book.value(PLAN_SHEET, r, md);
    const mdv = raw == null || raw === "" ? null : numeric(raw);
    if (mdv == null) continue;
    const n = plan.md.length;
    if (n > 0 && !(mdv > plan.md[n - 1]!)) continue;
    plan.md.push(mdv);
    plan.inc.push(vbaVal(book.value(PLAN_SHEET, r, inc)));
    plan.azi.push(vbaVal(book.value(PLAN_SHEET, r, azi)));
    plan.tvd.push(vbaVal(book.value(PLAN_SHEET, r, tvd)));
    plan.ns.push(vbaVal(book.value(PLAN_SHEET, r, ns)));
    plan.ew.push(vbaVal(book.value(PLAN_SHEET, r, ew)));
  }
  if (plan.md.length > 0) {
    const ne = integrateLocalNE(plan.md, plan.inc, plan.azi);
    plan.ns = ne.ns;
    plan.ew = ne.ew;
  }
  return plan;
}

function station(plan: Plan, i: number): PlanPoint {
  return { n: plan.ns[i]!, e: plan.ew[i]!, v: plan.tvd[i]!, azi: plan.azi[i]!, inc: plan.inc[i]! };
}

/** Direction at fraction f along the min-curvature arc from (i1,a1) to (i2,a2). */
function interpCourseDir(i1: number, a1: number, i2: number, a2: number, f: number): { inc: number; azi: number } {
  const n1 = Math.sin(rad(i1)) * Math.cos(rad(a1));
  const e1 = Math.sin(rad(i1)) * Math.sin(rad(a1));
  const v1 = Math.cos(rad(i1));
  const n2 = Math.sin(rad(i2)) * Math.cos(rad(a2));
  const e2 = Math.sin(rad(i2)) * Math.sin(rad(a2));
  const v2 = Math.cos(rad(i2));
  const beta = Math.acos(clamp(n1 * n2 + e1 * e2 + v1 * v2, -1, 1));
  let n: number;
  let e: number;
  let v: number;
  if (beta < 0.0000001) {
    n = n1;
    e = e1;
    v = v1;
  } else {
    const s = Math.sin(beta);
    n = (Math.sin((1 - f) * beta) * n1 + Math.sin(f * beta) * n2) / s;
    e = (Math.sin((1 - f) * beta) * e1 + Math.sin(f * beta) * e2) / s;
    v = (Math.sin((1 - f) * beta) * v1 + Math.sin(f * beta) * v2) / s;
  }
  const mag = Math.sqrt(n * n + e * e + v * v);
  if (mag > 0) {
    n /= mag;
    e /= mag;
    v /= mag;
  }
  const inc = Math.acos(clamp(v, -1, 1)) * 180 / Math.PI;
  if (Math.abs(n) < 1e-10 && Math.abs(e) < 1e-10) return { inc, azi: a1 };
  let azi = Math.atan2(e, n) * 180 / Math.PI;
  if (azi < 0) azi += 360;
  return { inc, azi };
}

/** PlanAt: min-curvature position on the plan at an MD, never a lerp of NS/EW/TVD. */
export function planAt(plan: Plan, md: number): PlanPoint {
  const n = plan.md.length;
  if (md <= plan.md[0]!) return station(plan, 0);
  if (md >= plan.md[n - 1]!) return station(plan, n - 1);
  let i = 0;
  for (i = 0; i <= n - 2; i += 1) if (md >= plan.md[i]! && md <= plan.md[i + 1]!) break;
  const span = plan.md[i + 1]! - plan.md[i]!;
  const f = span > 0 ? (md - plan.md[i]!) / span : 0;
  if (f <= 0) return station(plan, i);
  if (f >= 1) return station(plan, i + 1);
  const dir = interpCourseDir(plan.inc[i]!, plan.azi[i]!, plan.inc[i + 1]!, plan.azi[i + 1]!, f);
  const step = minCurveStep(plan.md[i]!, plan.inc[i]!, plan.azi[i]!, md, dir.inc, dir.azi);
  return { n: plan.ns[i]! + step.n, e: plan.ew[i]! + step.e, v: plan.tvd[i]! + step.tvd, azi: dir.azi, inc: dir.inc };
}

/** PlanAtTvd: bisect the first plan course that spans the TVD, else the nearest station. */
function planAtTvd(plan: Plan, tvd: number): PlanPoint {
  const n = plan.md.length;
  for (let i = 0; i <= n - 2; i += 1) {
    const tLo = plan.tvd[i]!;
    const tHi = plan.tvd[i + 1]!;
    if (Math.abs(tHi - tLo) < 0.0001) continue;
    if ((tvd >= tLo && tvd <= tHi) || (tvd >= tHi && tvd <= tLo)) {
      let lo = plan.md[i]!;
      let hi = plan.md[i + 1]!;
      const up = tHi >= tLo;
      for (let iter = 1; iter <= 28; iter += 1) {
        const mid = (lo + hi) / 2;
        const tMid = planAt(plan, mid).v;
        if ((up && tMid < tvd) || (!up && tMid > tvd)) lo = mid;
        else hi = mid;
      }
      return planAt(plan, (lo + hi) / 2);
    }
  }
  let best = 0;
  let bestD = Math.abs(plan.tvd[0]! - tvd);
  for (let i = 1; i < n; i += 1) {
    const d = Math.abs(plan.tvd[i]! - tvd);
    if (d < bestD) {
      bestD = d;
      best = i;
    }
  }
  return station(plan, best);
}

/** Linear waypoint TVD at MD from AC14:AD33; null outside the waypoint span or with < 2 points. */
export function waypointTvdAtMd(book: SheetBook, md: number): number | null {
  const mds: number[] = [];
  const tvds: number[] = [];
  for (let r = WP_FIRST; r <= WP_LAST; r += 1) {
    const vM = book.value(SS_SHEET, r, WP_COL_MD);
    const vT = book.value(SS_SHEET, r, WP_COL_TVD);
    const m = vM == null || vM === "" ? null : numeric(vM);
    const t = vT == null || vT === "" ? null : numeric(vT);
    if (m == null || t == null) continue;
    if (mds.length === 0 || m > mds[mds.length - 1]!) {
      mds.push(m);
      tvds.push(t);
    }
  }
  const n = mds.length;
  if (n < 2 || md < mds[0]! || md > mds[n - 1]!) return null;
  let i = 0;
  for (i = 0; i <= n - 2; i += 1) if (md >= mds[i]! && md <= mds[i + 1]!) break;
  if (i > n - 2) i = n - 2;
  const span = mds[i + 1]! - mds[i]!;
  const f = span > 0 ? (md - mds[i]!) / span : 0;
  return tvds[i]! + f * (tvds[i + 1]! - tvds[i]!);
}

export function waypointHalfWidth(book: SheetBook): number {
  const v = numeric(book.value(SS_SHEET, 14, 28));
  return v != null && v > 0 ? v : 0;
}

export function isSummaryRow(book: SheetBook, r: number): boolean {
  return [5, 6, 4].some((c) => (book.formula(SS_SHEET, r, c) ?? "").toUpperCase().includes("LOOKUP"));
}

function surveyStation(book: SheetBook, r: number): { md: number; inc: number; azi: number } | null {
  const md = book.value(SS_SHEET, r, 5);
  const inc = book.value(SS_SHEET, r, 6);
  const azi = book.value(SS_SHEET, r, 7);
  const values = [md, inc, azi].map((v) => (v == null || v === "" ? null : numeric(v)));
  if (values.some((v) => v == null)) return null;
  return { md: values[0]!, inc: values[1]!, azi: values[2]! };
}

function actualAtLastSurvey(book: SheetBook): { count: number; md: number; inc: number; azi: number; n: number; e: number; v: number; row: number } {
  let prev = { md: 0, inc: 0, azi: 0 };
  let n = 0;
  let e = 0;
  let v = 0;
  let row = 0;
  let count = 0;
  for (let r = SURV_FIRST; r <= SURV_LAST; r += 1) {
    if (isSummaryRow(book, r) || !isGoodSurveyRow(book, r)) continue;
    const st = surveyStation(book, r);
    if (!st || st.md <= prev.md) continue;
    const step = minCurveStep(prev.md, prev.inc, prev.azi, st.md, st.inc, st.azi);
    n += step.n;
    e += step.e;
    v += step.tvd;
    prev = st;
    row = r;
    count += 1;
  }
  return { count, ...prev, n, e, v, row };
}

/** Deepest numeric survey (E/F/G) with any fill: the station the user typed. */
function lastEnteredSurvey(book: SheetBook): { md: number; inc: number; azi: number; row: number } | null {
  let best: { md: number; inc: number; azi: number; row: number } | null = null;
  for (let r = SURV_FIRST; r <= SURV_LAST; r += 1) {
    if (isSummaryRow(book, r)) continue;
    const st = surveyStation(book, r);
    if (st && (!best || st.md > best.md)) best = { ...st, row: r };
  }
  return best;
}

export function frameComponents(gravity: boolean, planAzi: number, dN: number, dE: number, dTvdUp: number, planInc = 90): { x: number; y: number } {
  const right = dE * Math.cos(rad(planAzi)) - dN * Math.sin(rad(planAzi));
  const along = dN * Math.cos(rad(planAzi)) + dE * Math.sin(rad(planAzi));
  if (gravity) return { x: right, y: along * Math.cos(rad(planInc)) + dTvdUp * Math.sin(rad(planInc)) };
  return { x: right, y: along };
}

function planDist2(plan: Plan, q: number, n: number, e: number, v: number): number {
  const c = planAt(plan, q);
  return (n - c.n) ** 2 + (e - c.e) ** 2 + (v - c.v) ** 2;
}

/** Closest point on the plan to (n, e, v), searched within +-60 m MD of md0. */
function closestPlanPoint(plan: Plan, md0: number, n: number, e: number, v: number): PlanPoint {
  const lo = Math.max(md0 - 60, plan.md[0]!);
  const hi = Math.min(md0 + 60, plan.md[plan.md.length - 1]!);
  let bestD = -1;
  let bestQ = lo;
  for (let q = lo; q <= hi; q += 1) {
    const d = planDist2(plan, q, n, e, v);
    if (bestD < 0 || d < bestD) {
      bestD = d;
      bestQ = q;
    }
  }
  let a = Math.max(bestQ - 1, lo);
  let b = Math.min(bestQ + 1, hi);
  for (let k = 1; k <= 40; k += 1) {
    const m1 = a + (b - a) / 3;
    const m2 = b - (b - a) / 3;
    if (planDist2(plan, m1, n, e, v) < planDist2(plan, m2, n, e, v)) b = m2;
    else a = m1;
  }
  return planAt(plan, (a + b) / 2);
}

/** y = plan TVD - actual TVD there (UP +); x = horizontal distance, RT + / LT - of plan azimuth. */
function closestComponents(plan: Plan, md0: number, n: number, e: number, v: number): { x: number; y: number } {
  const c = closestPlanPoint(plan, md0, n, e, v);
  const dN = n - c.n;
  const dE = e - c.e;
  const rt = dE * Math.cos(rad(c.azi)) - dN * Math.sin(rad(c.azi));
  let x = Math.sqrt(dN * dN + dE * dE);
  if (rt < 0) x = -x;
  return { x, y: c.v - v };
}

function actualTvdForGauge(book: SheetBook, row: number, integrated: number): number {
  const h = book.value(SS_SHEET, row, 8);
  const n = h == null || h === "" ? null : numeric(h);
  if (n == null) return integrated;
  return Math.abs(n - integrated) > 1 ? integrated : n;
}

export function slideIntervals(book: SheetBook): SlideInterval[] {
  const out: SlideInterval[] = [];
  for (let r = 13; r <= 505; r += 1) {
    const from = numeric(book.value(SS_SHEET, r, 18));
    const to = numeric(book.value(SS_SHEET, r, 19));
    if (from == null || to == null) continue;
    const tf = parseToolface(book.value(SS_SHEET, r, 21));
    if (tf == null) continue;
    out.push({ from, to, tf, row: r - 12 });
  }
  return out;
}

export interface GaugeResult {
  model: GaugeModel;
  /** Inputs for UpdateWellborePosition when the dial had data. */
  wellbore: { lat: number; planUp: number; row: number; md: number; inc: number } | null;
}

const NO_DATA: GaugeModel = { hasData: false, caption: "NO DATA - Import Plan / enter surveys", mode: "GRAV", ax: 0, ay: 0, showPtb: false, bx: 0, by: 0, hasBand: false, yTop: 0, yBot: 0, scale: 1 };

export function computePlanGauge(book: SheetBook): GaugeResult {
  if (!book.hasSheet(SS_SHEET)) return { model: NO_DATA, wellbore: null };
  const plan = loadPlan(book);
  const surv = actualAtLastSurvey(book);
  if (!plan || plan.md.length < 2 || surv.count < 1) return { model: NO_DATA, wellbore: null };
  let sN = surv.n;
  let sE = surv.e;
  const vN = book.value(SS_SHEET, surv.row, 38);
  const vE = book.value(SS_SHEET, surv.row, 39);
  if (isNumber(vN) && isNumber(vE)) {
    sN = vN;
    sE = vE;
  }
  const actTvd = actualTvdForGauge(book, surv.row, surv.v);
  const gravity = surv.inc >= CROSSOVER_INC;
  const pl = gravity ? planAt(plan, surv.md) : planAtTvd(plan, actTvd);
  const halfW = waypointHalfWidth(book);
  const wpTvd = gravity && halfW > 0 ? waypointTvdAtMd(book, surv.md) : null;
  const sail = wpTvd != null;
  const closest = gravity && !sail;
  let a = frameComponents(gravity, pl.azi, sN - pl.n, sE - pl.e, pl.v - actTvd, sail ? 90 : pl.inc);
  if (closest) a = closestComponents(plan, surv.md, sN, sE, actTvd);

  let showPtb = false;
  let b = { x: 0, y: 0 };
  const entered = lastEnteredSurvey(book);
  if (entered && entered.row !== surv.row && entered.md > surv.md + 0.01) {
    const step = minCurveStep(surv.md, surv.inc, surv.azi, entered.md, entered.inc, entered.azi);
    const bitTvd = actTvd + step.tvd;
    const pe = gravity ? planAt(plan, entered.md) : planAtTvd(plan, bitTvd);
    b = frameComponents(gravity, pe.azi, sN + step.n - pe.n, sE + step.e - pe.e, pe.v - bitTvd, sail ? 90 : pe.inc);
    if (closest) b = closestComponents(plan, entered.md, sN + step.n, sE + step.e, bitTvd);
    showPtb = true;
  }
  const bitMd = numeric(book.value(SS_SHEET, surv.row, 4));
  if (!showPtb && bitMd != null && bitMd > surv.md + 0.01) {
    const vW = book.value(SS_SHEET, surv.row, 23);
    const vX = book.value(SS_SHEET, surv.row, 24);
    const incB = vW == null || vW === "" ? surv.inc : numeric(vW) ?? surv.inc;
    const azB = vX == null || vX === "" ? surv.azi : numeric(vX) ?? surv.azi;
    const motor = numeric(book.value(SS_SHEET, surv.row, 35)) ?? 0;
    const walked = walkUnseenToBit(surv.md, surv.inc, surv.azi, bitMd, motor, slideIntervals(book));
    const step = walked
      ? { tvd: walked.tvd, n: walked.n, e: walked.e }
      : minCurveStep(surv.md, surv.inc, surv.azi, bitMd, incB, azB);
    const bitTvd = actTvd + step.tvd;
    const pb = gravity ? planAt(plan, bitMd) : planAtTvd(plan, bitTvd);
    b = frameComponents(gravity, pb.azi, sN + step.n - pb.n, sE + step.e - pb.e, pb.v - bitTvd, sail ? 90 : pb.inc);
    if (closest) b = closestComponents(plan, bitMd, sN + step.n, sE + step.e, bitTvd);
    showPtb = true;
  }

  let hasBand = false;
  let yTop = 0;
  let yBot = 0;
  let scale = 0;
  if (sail) {
    const yWp = pl.v - wpTvd!;
    yTop = yWp + halfW;
    yBot = yWp - halfW;
    hasBand = true;
    scale = Math.max(halfW, Math.abs(yTop), Math.abs(yBot)) * 1.35;
    if (scale < 0.5) scale = 0.5;
  }
  const magA = Math.hypot(a.x, a.y);
  const magB = Math.hypot(b.x, b.y);
  if (!hasBand) {
    scale = halfW > 0 ? halfW : 1;
    if (magA > scale) scale = magA;
    if (scale < 0.5) scale = 0.5;
  }
  if (showPtb && magB > scale) scale = magB * 1.15;
  if (scale < 0.5) scale = 0.5;
  return {
    model: { hasData: true, caption: "", mode: gravity ? "GRAV" : "MAG", ax: a.x, ay: a.y, showPtb, bx: b.x, by: b.y, hasBand, yTop, yBot, scale },
    wellbore: { lat: a.x, planUp: a.y, row: surv.row, md: surv.md, inc: surv.inc },
  };
}

/** AxisLabel: "DN 6.06m", "RT 11.96m", "AH 0.40m". */
export function axisLabel(mode: GaugeMode, vertical: boolean, v: number): string {
  let tag: string;
  if (mode === "GRAV") tag = vertical ? (v >= 0 ? "UP" : "DN") : v >= 0 ? "RT" : "LT";
  else tag = vertical ? (v >= 0 ? "AH" : "BH") : v >= 0 ? "RT" : "LT";
  return `${`${tag}   `.slice(0, 3)}${fixed2(Math.abs(v))}m`;
}

function firstSailMd(book: SheetBook): number | null {
  for (let r = WP_FIRST; r <= WP_LAST; r += 1) {
    const v = book.value(SS_SHEET, r, WP_COL_MD);
    if (v != null && v !== "" && numeric(v) != null) return numeric(v);
  }
  return null;
}

function geoWindowAt(book: SheetBook, row: number): number | null {
  const read = (r: number) => {
    const v = book.value(SS_SHEET, r, 28);
    const text = cellText(v);
    if (text === "" || !isNumber(v) && numeric(v) == null) return null;
    return numeric(v);
  };
  if (row >= SURV_FIRST && row <= SURV_LAST && row !== 14) {
    const v = read(row);
    if (v != null) return v;
  }
  for (let r = SURV_LAST; r >= SURV_FIRST; r -= 1) {
    if (r === 14 || isSummaryRow(book, r) || !isGoodSurveyRow(book, r)) continue;
    if (!surveyStation(book, r)) continue;
    const v = read(r);
    if (v != null) return v;
  }
  return null;
}

/** UpdateWellborePosition: Data E13:F15. Only cells whose text changes are written. */
export function updateWellborePosition(book: SheetBook, wellbore: NonNullable<GaugeResult["wellbore"]>): void {
  if (!book.hasSheet("Data")) return;
  const { lat, planUp, row, md, inc } = wellbore;
  let geo: number | null;
  const sailMd = inc >= INC_LATERAL ? firstSailMd(book) : null;
  if (sailMd != null && md + 0.0001 >= sailMd) geo = geoWindowAt(book, row);
  else geo = Math.hypot(lat, planUp) * (planUp < 0 ? -1 : 1);
  const put = (r: number, c: number, text: string) => {
    if (cellText(book.value("Data", r, c)) !== text) book.setValue("Data", r, c, text);
  };
  put(13, 5, `${fixed2(Math.abs(lat))}m`);
  put(13, 6, lat >= 0 ? "Right" : "Left");
  put(14, 5, `${fixed2(Math.abs(planUp))}m`);
  put(14, 6, planUp < 0 ? "Below" : "Above");
  if (geo != null) {
    put(15, 5, `${fixed2(Math.abs(geo))}m`);
    put(15, 6, geo < 0 ? "Below" : "Above");
  }
}
