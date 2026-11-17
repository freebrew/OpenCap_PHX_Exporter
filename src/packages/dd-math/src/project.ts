import { arcToPoint, deg2rad, doglegAngleDeg, EPS, interpCourseDir, inverseCourse, minCurveStep, toolfaceBetweenDeg, wrap180, wrap360 } from "./angles";
import { planAt, type PlanStation } from "./plan";

const MIN_AIM_MD = 10;
const MIN_AIM_TVD = 5;
const MIN_AIM_DINC = 0.5;
const SLIDE_IGNORE_M = 1;
const AIM_WALK_STEP_M = 3;
const AIM_WALK_MAX_M = 90;
const CATCHUP_CURVE_MIN = 1;
const CATCHUP_LATERAL_INC = 80;

export class UdfError extends Error {
  readonly code: "VALUE" | "NUM";

  constructor(code: "VALUE" | "NUM") {
    super(code);
    this.name = "UdfError";
    this.code = code;
  }
}

export interface TargetStation {
  md: number;
  inc: number;
  azm: number;
  tvd: number;
}

export interface SlideInterval {
  from: number;
  to: number;
  tf: number;
  row: number;
}

export interface CallerSurvey {
  md: number;
  bitMd: number;
  inc: number;
  azm: number;
  tvd: number | null;
  cumN: number | null;
  cumE: number | null;
  motor: number | null;
  course?: number | null;
}

export interface SurveyRowLookup {
  cumN: number | null;
  cumE: number | null;
  motor: number | null;
}

export interface SurveyContext {
  caller?: CallerSurvey | null;
  slides?: SlideInterval[];
  plan?: PlanStation[];
  /** Slidesheet survey row whose MD (E) is within 0.01 of md, searched bottom-up. */
  surveyRow?: (md: number) => SurveyRowLookup | null;
}

export function num(value: unknown): number | null {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value === "boolean") return value ? 1 : 0;
  if (typeof value === "string") {
    const text = value.trim();
    if (text === "" || !Number.isFinite(Number(text))) return null;
    return Number(text);
  }
  return null;
}

export function column(value: unknown): unknown[] {
  if (!Array.isArray(value)) return [value];
  return value.map((row) => (Array.isArray(row) ? row[0] : row));
}

export function parseToolface(text: unknown): number | null {
  if (text instanceof UdfError) throw text;
  if (typeof text === "object" && text && "code" in text) throw new UdfError("VALUE");
  let source = String(text ?? "").replaceAll("\u00a0", " ").trim().toUpperCase().replaceAll(" ", "");
  if (source === "") return null;
  if (source.endsWith("M")) {
    const magnetic = source.slice(0, -1);
    if (magnetic !== "" && Number.isFinite(Number(magnetic))) return Number(magnetic);
    throw new UdfError("VALUE");
  }
  if (source === "R" || source === "HS" || source === "HSR") return 0;
  if (source === "L" || source === "LS" || source === "HSL") return 180;
  if ((source.startsWith("R") || source.startsWith("L")) && Number.isFinite(Number(source.slice(1)))) {
    const magnitude = Math.abs(Number(source.slice(1)));
    return source.startsWith("R") ? magnitude : -magnitude;
  }
  if (source.length > 1 && (source.endsWith("R") || source.endsWith("L")) && Number.isFinite(Number(source.slice(0, -1)))) {
    const magnitude = Math.abs(Number(source.slice(0, -1)));
    return source.endsWith("R") ? magnitude : -magnitude;
  }
  if (Number.isFinite(Number(source))) return Number(source);
  throw new UdfError("VALUE");
}

export function doglegBelow(dls: number, course: number, seen: number, below: number): number {
  if (Math.abs(seen) < EPS || Math.abs(below) < EPS) return 0;
  return dls / seen * course / 30 * below;
}

export function doglegBelowYield(motorOut: number, seen: number, below: number): number {
  if (motorOut <= EPS || seen + 0.0000001 < 2 || Math.abs(below) < EPS) return 0;
  return motorOut / 30 * below;
}

export function azmWalkDeg(survInc: number, dogleg: number, toolface: number): number {
  if (Math.abs(dogleg) < EPS) return 0;
  const sine = Math.sin(deg2rad(survInc));
  if (survInc >= 5 && Math.abs(sine) > EPS) return dogleg * Math.sin(deg2rad(toolface)) / sine;
  return dogleg * Math.sin(deg2rad(toolface));
}

function signedSeenTf(seenTf: unknown, seenLr: unknown): number {
  const value = num(seenTf);
  if (value == null) return 0;
  const side = String(seenLr ?? "").trim().toUpperCase();
  if (side === "L") return -Math.abs(value);
  if (side === "R") return Math.abs(value);
  return 0;
}

export function resolveWalkTf(dial: unknown, seenTf: unknown, seenLr: unknown): number {
  const dialTf = num(dial) ?? 0;
  const effective = signedSeenTf(seenTf, seenLr);
  if (Math.abs(effective) < 0.05) return dialTf;
  if (dialTf * effective > 0 && Math.abs(effective) < Math.abs(dialTf)) return effective;
  return dialTf;
}

export function walkUnseenToBit(survMd: number, survInc: number, survAzm: number, bitMd: number, motorOut: number, slides: SlideInterval[]): { inc: number; azm: number; tvd: number; n: number; e: number } | null {
  if (bitMd <= survMd + 0.01) return null;
  const clipped = slides.flatMap((slide) => {
    let from = slide.from;
    let to = slide.to;
    if (!(to > from + 0.001)) return [];
    if (from < survMd) from = survMd;
    if (to > bitMd) to = bitMd;
    if (!(to > from + 0.001) || !Number.isFinite(slide.tf)) return [];
    return [{ from, to, tf: slide.tf, row: slide.row }];
  });
  const breaks = [survMd, bitMd, ...clipped.flatMap((slide) => [slide.from, slide.to])].sort((a, b) => a - b);
  const unique = breaks.filter((value, index) => index === 0 || value - (breaks[index - 1] ?? value) > 0.0005);
  let inc = survInc;
  let azm = survAzm;
  let tvd = 0;
  let n = 0;
  let e = 0;
  for (let i = 0; i < unique.length - 1; i += 1) {
    const md1 = unique[i] ?? survMd;
    const md2 = unique[i + 1] ?? bitMd;
    if (md2 - md1 < 0.001) continue;
    const mid = (md1 + md2) / 2;
    let bestRow = 0;
    let tf = 0;
    for (const slide of clipped) {
      if (slide.from <= mid + 0.0005 && mid < slide.to - 0.0000001 && slide.row >= bestRow) {
        bestRow = slide.row;
        tf = slide.tf;
      }
    }
    let inc2 = inc;
    let azm2 = azm;
    if (bestRow > 0 && motorOut > EPS) {
      const dogleg = motorOut / 30 * (md2 - md1);
      inc2 = inc + dogleg * Math.cos(deg2rad(tf));
      azm2 = wrap360(azm + azmWalkDeg(inc, dogleg, tf));
    }
    if (inc2 < 0) inc2 = 0;
    if (inc2 > 180) inc2 = 180;
    const step = minCurveStep(md1, inc, azm, md2, inc2, azm2);
    tvd += step.tvd;
    n += step.n;
    e += step.e;
    inc = inc2;
    azm = azm2;
  }
  return { inc, azm, tvd, n, e };
}

export function loadTargets(md: unknown, inc: unknown, azm: unknown, tvd: unknown): TargetStation[] {
  const mds = column(md);
  const incs = column(inc);
  const azms = column(azm);
  const tvds = column(tvd);
  const targets: TargetStation[] = [];
  for (let i = 0; i < mds.length && targets.length < 100; i += 1) {
    const stationMd = num(mds[i]);
    if (stationMd == null) continue;
    targets.push({
      md: stationMd,
      inc: num(incs[i]) ?? 0,
      azm: num(azms[i]) ?? 0,
      tvd: num(tvds[i]) ?? 0,
    });
  }
  return targets;
}

function targetExhausted(bitMd: number, incBit: number, target: TargetStation, tvdBit: number | null): boolean {
  if (incBit + MIN_AIM_DINC < target.inc) return false;
  const distance = target.md - bitMd;
  if (distance <= EPS || distance < MIN_AIM_MD) return true;
  if (tvdBit != null) {
    const dTvd = target.tvd - tvdBit;
    if (dTvd > 0 && dTvd < MIN_AIM_TVD) return true;
  }
  return false;
}

function pastFinal(inc: number, tvd: number | null, targets: TargetStation[]): boolean {
  const last = targets[targets.length - 1];
  if (!last) return false;
  if (inc >= last.inc) return true;
  return tvd != null && tvd >= last.tvd;
}

export function firstAimIndex(targets: TargetStation[], bitMd: number, incBit: number, tvdBit: number | null): number {
  let index = targets.findIndex((target) => target.md > bitMd);
  // VBA's `Do While k <= n And m(k) <= bitMd` does not short-circuit: with the bit past
  // every target it reads m(n + 1), and each public caller's handler returns #NUM!.
  if (index < 0) throw new UdfError("NUM");
  while (index < targets.length) {
    const target = targets[index];
    if (!target || index === targets.length - 1) return index;
    if (!targetExhausted(bitMd, incBit, target, tvdBit)) return index;
    index += 1;
  }
  return targets.length;
}

function planAttitude(targets: TargetStation[], md: number): { inc: number; azm: number } {
  const first = targets[0];
  const last = targets[targets.length - 1];
  if (!first || !last) return { inc: 0, azm: 0 };
  if (targets.length < 2 || md <= first.md) return { inc: first.inc, azm: first.azm };
  if (md >= last.md) return { inc: last.inc, azm: last.azm };
  let index = 1;
  while (index < targets.length && (targets[index]?.md ?? md) < md) index += 1;
  const right = targets[Math.min(index, targets.length - 1)];
  const left = targets[Math.max(0, Math.min(index, targets.length - 1) - 1)];
  if (!left || !right || Math.abs(right.md - left.md) < EPS) return { inc: right?.inc ?? 0, azm: right?.azm ?? 0 };
  const fraction = (md - left.md) / (right.md - left.md);
  return interpCourseDir(left.inc, left.azm, right.inc, right.azm, fraction);
}

export interface Aim {
  distance: number;
  dls: number;
  burr: number;
  tf: number;
  inc: number;
  tvd: number;
}

export function computeAim(bitMd: number, incBit: number, azmBit: number, motorOut: number, course: number, targets: TargetStation[], tvdBit: number | null, allowWalk = true): Aim | null {
  if (targets.length < 1) return null;
  const index = firstAimIndex(targets, bitMd, incBit, tvdBit);
  const aim = targets[index];
  if (!aim) return null;
  let md = aim.md;
  let distance = md - bitMd;
  let attitude = { inc: aim.inc, azm: aim.azm };
  if (distance < course && incBit + MIN_AIM_DINC < attitude.inc && course > EPS) distance = course;
  if (distance <= EPS) return null;
  let dls = doglegAngleDeg(incBit, azmBit, attitude.inc, attitude.azm) * 30 / distance;
  const last = targets[targets.length - 1];
  if (allowWalk && motorOut > EPS && course > EPS && last) {
    while (md < last.md && dls > motorOut + EPS) {
      md += course;
      if (md > last.md) md = last.md;
      attitude = planAttitude(targets, md);
      distance = md - bitMd;
      if (distance <= EPS) break;
      dls = doglegAngleDeg(incBit, azmBit, attitude.inc, attitude.azm) * 30 / distance;
      if (Math.abs(md - last.md) < EPS && dls > motorOut) break;
    }
  }
  if (distance <= EPS) return null;
  return {
    distance,
    dls,
    burr: (attitude.inc - incBit) * 30 / distance,
    tf: toolfaceBetweenDeg(incBit, azmBit, attitude.inc, attitude.azm),
    inc: attitude.inc,
    tvd: aim.tvd,
  };
}

export function aimPlanPoint(bitMd: number, incBit: number, azmBit: number, north: number, east: number, tvd: number, motorOut: number, targets: TargetStation[], plan: PlanStation[], fixedMd: number | null = null): { tf: number; dls: number; aimMd: number } | null {
  if (plan.length < 2) return null;
  let md = fixedMd ?? 0;
  if (fixedMd == null) {
    if (targets.length < 1) return null;
    const index = firstAimIndex(targets, bitMd, incBit, tvd);
    const target = targets[index];
    if (!target) return null;
    md = target.md;
  } else if (md < bitMd + AIM_WALK_STEP_M) {
    md = bitMd + AIM_WALK_STEP_M;
  }
  const end = plan[plan.length - 1]?.md ?? md;
  if (md > end) md = end;
  const walkEnd = Math.min(end, md + AIM_WALK_MAX_M);
  let cursor = 0;
  while (cursor < plan.length - 2 && (plan[cursor + 1]?.md ?? bitMd) <= bitMd) cursor += 1;
  let planMax = 0;
  let result: { tf: number; dls: number } | null = null;
  for (;;) {
    const point = planAt(md, plan);
    if (!point) return null;
    let pointTvd = point.tvd;
    for (const target of targets) {
      if (Math.abs(target.md - md) < 0.05) {
        if (Math.abs(target.tvd - point.tvd) > 0.5) pointTvd = target.tvd;
        break;
      }
    }
    result = arcToPoint(incBit, azmBit, point.n - north, point.e - east, pointTvd - tvd);
    if (!result) return null;
    while (cursor <= plan.length - 2) {
      const station = plan[cursor];
      const next = plan[cursor + 1];
      if (!station || !next || station.md >= md) break;
      if (next.md - station.md > EPS) {
        const rate = doglegAngleDeg(station.inc, station.azm, next.inc, next.azm) * 30 / (next.md - station.md);
        if (rate > planMax) planMax = rate;
      }
      cursor += 1;
    }
    let cap = planMax >= CATCHUP_CURVE_MIN ? planMax : motorOut;
    if (planMax >= CATCHUP_CURVE_MIN && motorOut > EPS && motorOut < cap) cap = motorOut;
    if (cap <= EPS || result.dls <= cap + EPS || md >= walkEnd - EPS) break;
    md = Math.min(walkEnd, md + AIM_WALK_STEP_M);
  }
  if (!result) return null;
  return { ...result, aimMd: md };
}

/** WorksheetFunction.Round: halves go away from zero (Math.round sends -2.5 to -2). */
function excelRound(value: number): number {
  return Math.sign(value) * Math.round(Math.abs(value));
}

export function formatRequiredTf(tf: number, incBit: number, azmBit: number, threshold: number | null, tfText = ""): string {
  const mode = toolfaceMode(tfText, incBit, threshold);
  if (mode === "Magnetic") return `${excelRound(wrap360(azmBit + tf))}M`;
  const rounded = excelRound(wrap180(tf));
  if (rounded === 0) return "0";
  return rounded > 0 ? `R${rounded}` : `L${Math.abs(rounded)}`;
}

export function toolfaceMode(tfText: unknown, incBit: unknown, threshold: unknown): "Magnetic" | "Gravity" {
  const limit = num(threshold);
  const cutoff = limit != null && limit > 0 ? limit : 5;
  const source = String(tfText ?? "").replaceAll("\u00a0", " ").trim().toUpperCase();
  if (source.endsWith("M")) return "Magnetic";
  if (["R", "L", "HS", "LS", "HSR", "HSL"].includes(source)) return "Gravity";
  if ((source.startsWith("R") || source.startsWith("L")) && Number.isFinite(Number(source.slice(1)))) return "Gravity";
  if (source.length > 1 && (source.endsWith("R") || source.endsWith("L")) && Number.isFinite(Number(source.slice(0, -1)))) return "Gravity";
  const inc = num(incBit) ?? 0;
  return inc < cutoff ? "Magnetic" : "Gravity";
}

export function slideMetersBetween(fromMd: unknown, toMd: unknown, fromColumn: unknown, toColumn: unknown): number {
  const from = num(fromMd);
  const to = num(toMd);
  if (from == null || to == null || to <= from) return 0;
  const starts = column(fromColumn);
  const ends = column(toColumn);
  let total = 0;
  for (let i = 0; i < starts.length; i += 1) {
    const a0 = num(starts[i]);
    const b0 = num(ends[i]);
    if (a0 == null || b0 == null) continue;
    const a = Math.max(a0, from);
    const b = Math.min(b0, to);
    if (b > a) total += b - a;
  }
  return total;
}

export function rollingMotorOut(quality: unknown, metres: unknown): number | null {
  const q = column(quality);
  const m = column(metres);
  if (q.length < 1 || m.length !== q.length) return null;
  let taken = 0;
  let total = 0;
  let weight = 0;
  let thisQ: number | null = null;
  for (let i = q.length - 1; i >= 0 && taken < 3; i -= 1) {
    const value = num(q[i]);
    if (value == null || value <= 0) continue;
    const seen = num(m[i]);
    const sampleWeight = seen != null && seen > 0 ? seen : 1;
    total += value * sampleWeight;
    weight += sampleWeight;
    taken += 1;
    if (i === q.length - 1) thisQ = value;
  }
  if (taken >= 2 && weight > 0) return total / weight;
  return thisQ;
}

export function activeTargetMd(bitMd: unknown, targets: unknown): number | null {
  const bit = num(bitMd);
  if (bit == null) return null;
  for (const value of column(targets)) {
    const md = num(value);
    if (md != null && md > bit) return md;
  }
  return null;
}

export function activeTargetMdEx(bitMd: unknown, incBit: unknown, tvdBit: unknown, md: unknown, inc: unknown, azm: unknown, tvd: unknown, survInc: unknown, survTvd: unknown): number | null {
  const bit = num(bitMd);
  if (bit == null) return null;
  const targets = loadTargets(md, inc, azm, tvd);
  if (targets.length < 1) return null;
  const gateInc = num(survInc) ?? num(incBit) ?? 0;
  const gateTvd = num(survTvd) ?? num(tvdBit);
  if (pastFinal(gateInc, gateTvd, targets)) return null;
  const index = firstAimIndex(targets, bit, num(incBit) ?? 0, num(tvdBit));
  return targets[index]?.md ?? null;
}

export function metersToSlide(bitMd: unknown, incBit: unknown, azmBit: unknown, motorOut: unknown, course: unknown, md: unknown, inc: unknown, azm: unknown, tvd: unknown, tvdBit: unknown, context?: SurveyContext): number | null {
  const bm = num(bitMd);
  const ib = num(incBit);
  const ab = num(azmBit);
  const motor = num(motorOut) ?? 0;
  const stand = num(course) ?? 0;
  if (bm == null || ib == null || ab == null || motor <= EPS || stand <= EPS) return null;
  const targets = loadTargets(md, inc, azm, tvd);
  const aim = computeAim(bm, ib, ab, motor, stand, targets, num(tvdBit), false);
  if (!aim) return null;
  let required = aim.dls;
  const planned = planPointForCaller(bm, ib, ab, motor, num(tvdBit), targets, context);
  if (planned) required = planned.dls;
  else {
    const catchUp = catchUpDls(bm, ib, ab, stand, context?.plan);
    if (catchUp && catchUp.dls > EPS) required = catchUp.dls;
  }
  if (Math.abs(required) <= EPS) return 0;
  let metres = Math.abs(required) / motor * stand;
  if (metres > stand) metres = stand;
  if (metres < SLIDE_IGNORE_M) return 0;
  return metres;
}

function planPointForCaller(bitMd: number, incBit: number, azmBit: number, motorOut: number, tvdBit: number | null, targets: TargetStation[], context: SurveyContext | undefined): { tf: number; dls: number; aimMd: number } | null {
  const caller = context?.caller;
  const plan = context?.plan;
  if (!caller || !plan || caller.cumN == null || caller.cumE == null) return null;
  const step = minCurveStep(caller.md, caller.inc, caller.azm, bitMd, incBit, azmBit);
  const tvd = tvdBit ?? (caller.tvd == null ? null : caller.tvd + step.tvd);
  if (tvd == null) return null;
  return aimPlanPoint(bitMd, incBit, azmBit, caller.cumN + step.n, caller.cumE + step.e, tvd, motorOut, targets, plan);
}

export function requiredTf(bitMd: unknown, incBit: unknown, azmBit: unknown, md: unknown, inc: unknown, azm: unknown, tvd: unknown, threshold: unknown, tvdBit: unknown, context?: SurveyContext): string | null {
  const bm = num(bitMd);
  const ib = num(incBit);
  const ab = num(azmBit);
  if (bm == null || ib == null || ab == null) return null;
  const targets = loadTargets(md, inc, azm, tvd);
  const index = firstAimIndex(targets, bm, ib, num(tvdBit));
  const aim = targets[index];
  if (!aim) return null;
  const planned = planPointForCaller(bm, ib, ab, context?.caller?.motor ?? 0, num(tvdBit), targets, context);
  if (planned) return formatRequiredTf(planned.tf, ib, ab, num(threshold));
  let tf = doglegAngleDeg(ib, ab, aim.inc, aim.azm) < 0.000001 ? 0 : toolfaceBetweenDeg(ib, ab, aim.inc, aim.azm);
  const course = context?.caller?.course ?? 0;
  if (course > EPS) {
    const catchUp = catchUpDls(bm, ib, ab, course, context?.plan);
    const sectionAim = catchUp && catchUp.dls > EPS ? computeAim(bm, ib, ab, 0, course, targets, num(tvdBit), false) : null;
    if (catchUp && sectionAim && catchUp.dls > Math.abs(sectionAim.dls)) tf = catchUp.tf;
  }
  return formatRequiredTf(tf, ib, ab, num(threshold));
}

/**
 * Plan dogleg rate over the next course and the toolface toward the plan attitude at its end.
 * Only on a curve (> 1 deg/30 m) below 80 deg; otherwise null.
 */
function catchUpDls(bitMd: number, incBit: number, azmBit: number, course: number, plan: PlanStation[] | undefined): { dls: number; tf: number } | null {
  if (course <= EPS || !plan || plan.length < 2 || bitMd < (plan[0]?.md ?? 0)) return null;
  const here = planAttitude(plan, bitMd);
  const ahead = planAttitude(plan, bitMd + course);
  if (ahead.inc >= CATCHUP_LATERAL_INC) return null;
  const dls = doglegAngleDeg(here.inc, here.azm, ahead.inc, ahead.azm) * 30 / course;
  if (dls <= CATCHUP_CURVE_MIN) return null;
  return { dls, tf: toolfaceBetweenDeg(incBit, azmBit, ahead.inc, ahead.azm) };
}

export function burr(bitMd: unknown, incBit: unknown, azmBit: unknown, motorOut: unknown, course: unknown, md: unknown, inc: unknown, azm: unknown, tvd: unknown, tvdBit: unknown, context?: SurveyContext): number | null {
  const bm = num(bitMd);
  const ib = num(incBit);
  const ab = num(azmBit);
  if (bm == null || ib == null || ab == null) return null;
  const targets = loadTargets(md, inc, azm, tvd);
  const aim = computeAim(bm, ib, ab, num(motorOut) ?? 0, num(course) ?? 0, targets, num(tvdBit), false);
  if (!aim) return null;
  const planned = planPointForCaller(bm, ib, ab, num(motorOut) ?? 0, num(tvdBit), targets, context);
  return planned ? planned.dls : aim.dls;
}

export function tfToTarget(incBit: unknown, azmBit: unknown, tgtInc: unknown, tgtAzm: unknown, threshold: unknown): string | null {
  const ib = num(incBit);
  const ab = num(azmBit);
  const ti = num(tgtInc);
  const ta = num(tgtAzm);
  if (ib == null || ab == null || ti == null || ta == null) return null;
  const tf = doglegAngleDeg(ib, ab, ti, ta) < 0.000001 ? 0 : toolfaceBetweenDeg(ib, ab, ti, ta);
  return formatRequiredTf(tf, ib, ab, num(threshold));
}

/** Circular-arc build rate that lands on a landing station's TVD; otherwise ProjBurr with no motor. */
export function landingBurr(bitMd: unknown, incBit: unknown, azmBit: unknown, tvdBit: unknown, course: unknown, md: unknown, inc: unknown, azm: unknown, tvd: unknown, landInc: unknown, context?: SurveyContext): number | null {
  const bm = num(bitMd);
  const ib = num(incBit);
  const tb = num(tvdBit);
  if (bm == null || ib == null || tb == null) return null;
  const landing = num(landInc) ?? 88;
  const targets = loadTargets(md, inc, azm, tvd);
  if (targets.length < 1) return null;
  const aim = targets[firstAimIndex(targets, bm, ib, tb)];
  if (!aim) return null;
  if (aim.inc >= landing) {
    const dTvd = aim.tvd - tb;
    const dSin = Math.sin(deg2rad(aim.inc)) - Math.sin(deg2rad(ib));
    if (dTvd > EPS && dSin > EPS) return (30 / (dTvd / dSin)) * 180 / Math.PI;
  }
  return burr(bitMd, incBit, azmBit, 0, course, md, inc, azm, tvd, tvdBit, context);
}

/** Cont DI plan-point aim from the Cont DI bit, tied to the Slidesheet survey row at survMd. */
export function contDiAim(part: unknown, survMd: unknown, survInc: unknown, survAzm: unknown, bitMd: unknown, incBit: unknown, azmBit: unknown, tvdBit: unknown, md: unknown, inc: unknown, azm: unknown, tvd: unknown, aimTgtMd: unknown, context?: SurveyContext): number | string | null {
  try {
    const sm = num(survMd);
    const si = num(survInc);
    const sa = num(survAzm);
    const bm = num(bitMd);
    const ib = num(incBit);
    const ab = num(azmBit);
    const tb = num(tvdBit);
    if (sm == null || si == null || sa == null || bm == null || ib == null || ab == null || tb == null) return null;
    const row = context?.surveyRow?.(sm);
    if (!row || row.cumN == null || row.cumE == null) return null;
    const step = minCurveStep(sm, si, sa, bm, ib, ab);
    const aim = aimPlanPoint(bm, ib, ab, row.cumN + step.n, row.cumE + step.e, tb, row.motor ?? 0, loadTargets(md, inc, azm, tvd), context?.plan ?? [], num(aimTgtMd));
    if (!aim) return null;
    const key = String(part ?? "").trim().toUpperCase();
    switch (key) {
      case "TF": return formatRequiredTf(aim.tf, ib, ab, 5);
      case "DLR": return aim.dls;
      case "BRR": return aim.dls * Math.cos(deg2rad(aim.tf));
      case "TRR":
      case "TRRS": {
        const turn = Math.sin(deg2rad(ib)) > 0.01 ? aim.dls * Math.sin(deg2rad(aim.tf)) / Math.sin(deg2rad(ib)) : 0;
        return key === "TRR" ? Math.abs(turn) : turn;
      }
      case "AIM": return aim.aimMd;
      default: return null;
    }
  } catch {
    // VBA's handler returns "" for every failure here, unlike the other aim functions.
    return null;
  }
}

export function cumulative(previous: unknown, md1: unknown, inc1: unknown, azm1: unknown, md2: unknown, inc2: unknown, azm2: unknown, isFirst: unknown, axis: "n" | "e"): number | null {
  if (num(md2) == null || num(inc2) == null || num(azm2) == null) return null;
  if (isFirst === true || isFirst === 1 || String(isFirst).toUpperCase() === "TRUE") return 0;
  const startMd = num(md1);
  const startInc = num(inc1);
  const startAzm = num(azm1);
  if (startMd == null || startInc == null || startAzm == null) return null;
  return (num(previous) ?? 0) + minCurveStep(startMd, startInc, startAzm, num(md2) ?? 0, num(inc2) ?? 0, num(azm2) ?? 0)[axis];
}

export function incAtBit(survInc: unknown, dls: unknown, course: unknown, seen: unknown, below: unknown, tf: unknown, motorOut: unknown, context?: SurveyContext): number | null {
  const inc = num(survInc);
  if (inc == null) return null;
  if (isArgError(tf)) throw new UdfError("VALUE");
  const walked = walkFromContext(num(motorOut) ?? num(dls) ?? 0, context);
  if (walked) return walked.inc;
  if (num(tf) == null && typeof tf !== "number") return inc;
  const dogleg = num(motorOut) != null
    ? doglegBelowYield(num(motorOut) ?? 0, num(seen) ?? 0, num(below) ?? 0)
    : doglegBelow(num(dls) ?? 0, num(course) ?? 0, num(seen) ?? 0, num(below) ?? 0);
  return inc + dogleg * Math.cos(deg2rad(num(tf) ?? 0));
}

export function azmAtBit(survInc: unknown, survAzm: unknown, dls: unknown, course: unknown, seen: unknown, below: unknown, tf: unknown, seenTf: unknown, seenLr: unknown, motorOut: unknown, context?: SurveyContext): number | null {
  const inc = num(survInc);
  const azm = num(survAzm);
  if (inc == null || azm == null || isArgError(tf)) return null;
  const walked = walkFromContext(num(motorOut) ?? num(dls) ?? 0, context);
  if (walked) return walked.azm;
  if (num(tf) == null) return azm;
  const dogleg = num(motorOut) != null
    ? doglegBelowYield(num(motorOut) ?? 0, num(seen) ?? 0, num(below) ?? 0)
    : doglegBelow(num(dls) ?? 0, num(course) ?? 0, num(seen) ?? 0, num(below) ?? 0);
  return wrap360(azm + azmWalkDeg(inc, dogleg, resolveWalkTf(tf, seenTf, seenLr)));
}

export function tvdAtBit(survMd: unknown, survInc: unknown, survAzm: unknown, survTvd: unknown, bitMd: unknown, incBit: unknown, azmBit: unknown, context?: SurveyContext): number | null {
  const md = num(survMd);
  const inc = num(survInc);
  const azm = num(survAzm);
  const tvd = num(survTvd);
  const bit = num(bitMd);
  const ib = num(incBit);
  const ab = num(azmBit);
  if (md == null || inc == null || azm == null || tvd == null || bit == null || ib == null || ab == null) return null;
  const walked = walkFromContext(0, context);
  if (walked) return tvd + walked.tvd;
  return tvd + minCurveStep(md, inc, azm, bit, ib, ab).tvd;
}

function walkFromContext(motorOut: number, context: SurveyContext | undefined): ReturnType<typeof walkUnseenToBit> {
  const caller = context?.caller;
  const slides = context?.slides ?? [];
  if (!caller) return null;
  const motor = motorOut > EPS ? motorOut : (caller.motor ?? 0);
  return walkUnseenToBit(caller.md, caller.inc, caller.azm, caller.bitMd, motor, slides);
}

export function cumDelta(isFirst: unknown, previous: unknown, delta: number | null): number | null {
  if (truthy(isFirst)) return 0;
  const prior = num(previous);
  if (prior == null || delta == null) throw new UdfError("NUM");
  return prior + delta;
}

function truthy(value: unknown): boolean {
  return value === true || value === 1 || String(value).toUpperCase() === "TRUE";
}

export function bitNorth(survMd: unknown, survInc: unknown, survAzm: unknown, cumN: unknown, bitMd: unknown, incBit: unknown, azmBit: unknown): number | null {
  return bitOffset(survMd, survInc, survAzm, cumN, bitMd, incBit, azmBit, "n");
}

export function bitEast(survMd: unknown, survInc: unknown, survAzm: unknown, cumE: unknown, bitMd: unknown, incBit: unknown, azmBit: unknown): number | null {
  return bitOffset(survMd, survInc, survAzm, cumE, bitMd, incBit, azmBit, "e");
}

function bitOffset(survMd: unknown, survInc: unknown, survAzm: unknown, cum: unknown, bitMd: unknown, incBit: unknown, azmBit: unknown, axis: "n" | "e"): number | null {
  const md = num(survMd);
  const inc = num(survInc);
  const azm = num(survAzm);
  const total = num(cum);
  const bit = num(bitMd);
  const ib = num(incBit);
  const ab = num(azmBit);
  if (md == null || inc == null || azm == null || total == null || bit == null || ib == null || ab == null) return null;
  const step = minCurveStep(md, inc, azm, bit, ib, ab);
  return total + step[axis];
}

export function contDiTvdAtBit(bitMd: unknown, bitInc: unknown, bitAzm: unknown, tieMd: unknown, tieInc: unknown, tieAzm: unknown, tieTvd: unknown, pdOffset: unknown, diMd: unknown, diInc: unknown, diAzm: unknown, pdInc: unknown, pdAzm: unknown): number | null {
  const bit = num(bitMd);
  const tie = num(tieMd);
  const inc0 = num(tieInc);
  const azm0 = num(tieAzm);
  const tvd0 = num(tieTvd);
  const ib = num(bitInc);
  const ab = num(bitAzm);
  if (bit == null || tie == null || inc0 == null || azm0 == null || tvd0 == null || ib == null || ab == null || bit <= tie) return null;
  let lastMd = tie;
  let lastInc = inc0;
  let lastAzm = azm0;
  let tvd = tvd0;
  const mds = column(diMd);
  const incs = column(diInc);
  const azms = column(diAzm);
  for (let i = 0; i < mds.length; i += 1) {
    const md = num(mds[i]);
    const inc = num(incs[i]);
    const azm = num(azms[i]);
    if (md == null || inc == null || !(md > lastMd && md < bit)) continue;
    const step = minCurveStep(lastMd, lastInc, lastAzm, md, inc, azm ?? lastAzm);
    tvd += step.tvd;
    lastMd = md;
    lastInc = inc;
    lastAzm = azm ?? lastAzm;
  }
  const offset = num(pdOffset);
  const probeInc = num(pdInc);
  if (offset != null && probeInc != null) {
    const probeMd = bit - offset;
    if (probeMd > lastMd && probeMd < bit) {
      const step = minCurveStep(lastMd, lastInc, lastAzm, probeMd, probeInc, num(pdAzm) ?? lastAzm);
      tvd += step.tvd;
      lastMd = probeMd;
      lastInc = probeInc;
      lastAzm = num(pdAzm) ?? lastAzm;
    }
  }
  return tvd + minCurveStep(lastMd, lastInc, lastAzm, bit, ib, ab).tvd;
}

const SURVEY_BIT_MD = 0;
const SURVEY_MD = 1;
const SURVEY_INC = 2;
const SURVEY_AZM = 3;
const SURVEY_TVD = 4;
const SURVEY_BIT_INC = 19;
const SURVEY_BIT_AZM = 20;

export function tdActualBit(field: unknown, surveyBlock: unknown): number | null {
  const rows = Array.isArray(surveyBlock) ? surveyBlock : [];
  let prevMd = 0;
  let prevInc = 0;
  let prevAzm = 0;
  let north = 0;
  let east = 0;
  let tvd = 0;
  let count = 0;
  let bitMd = 0;
  let bitInc = 0;
  let bitAzm = 0;
  let surveyTvd: number | null = null;
  for (const row of rows) {
    if (!Array.isArray(row)) continue;
    const md = num(row[SURVEY_MD]);
    const inc = num(row[SURVEY_INC]);
    const azm = num(row[SURVEY_AZM]);
    if (md == null || inc == null || azm == null || md <= prevMd) continue;
    const step = minCurveStep(prevMd, prevInc, prevAzm, md, inc, azm);
    north += step.n;
    east += step.e;
    tvd += step.tvd;
    prevMd = md;
    prevInc = inc;
    prevAzm = azm;
    bitMd = num(row[SURVEY_BIT_MD]) ?? md;
    bitInc = num(row[SURVEY_BIT_INC]) ?? inc;
    bitAzm = num(row[SURVEY_BIT_AZM]) ?? azm;
    surveyTvd = num(row[SURVEY_TVD]);
    count += 1;
  }
  if (count < 1) return null;
  if (surveyTvd != null) tvd = surveyTvd;
  let md = prevMd;
  let inc = prevInc;
  let azm = prevAzm;
  if (bitMd > prevMd + 0.005) {
    const step = minCurveStep(prevMd, prevInc, prevAzm, bitMd, bitInc, bitAzm);
    north += step.n;
    east += step.e;
    tvd += step.tvd;
    md = bitMd;
    inc = bitInc;
    azm = bitAzm;
  } else if (bitMd > 0) {
    md = bitMd;
    inc = bitInc;
    azm = bitAzm;
  }
  return tdField(field, { md, inc, azm, n: north, e: east, tvd });
}

/** Planned TD MD plus (bit → TD inverse course) minus (plan at the bit's MD → TD inverse course). */
export function tdEstFinal(field: unknown, surveyBlock: unknown, plannedBlock: unknown, plan: PlanStation[] = []): number | null {
  const actual = tdVector(surveyBlock);
  const planned = plannedVector(plannedBlock);
  if (!actual || !planned) return null;
  const usable = plan.length >= 2;
  let plannedTvd = usable ? planAt(planned.md, plan)?.tvd ?? 0 : 0;
  if (plannedTvd === 0 && actual.tvd !== 0) plannedTvd = actual.tvd;
  const toTd = (inc: number, azm: number, n: number, e: number, tvd: number) =>
    inverseCourse(inc, azm, n, e, tvd, planned.inc, planned.azm, planned.n, planned.e, plannedTvd);
  const actualInverse = toTd(actual.inc, actual.azm, actual.n, actual.e, actual.tvd);
  const onPlan = usable ? planAt(actual.md, plan) : null;
  let md = onPlan
    ? planned.md + (actualInverse - toTd(onPlan.inc, onPlan.azm, onPlan.n, onPlan.e, onPlan.tvd))
    : actual.md + actualInverse;
  if (md < actual.md) md = actual.md;
  return tdField(field, { ...planned, md, tvd: plannedTvd });
}

function tdVector(surveyBlock: unknown): { md: number; inc: number; azm: number; n: number; e: number; tvd: number } | null {
  const md = tdActualBit("MD", surveyBlock);
  const inc = tdActualBit("INC", surveyBlock);
  const azm = tdActualBit("AZM", surveyBlock);
  const n = tdActualBit("NS", surveyBlock);
  const e = tdActualBit("EW", surveyBlock);
  const tvd = tdActualBit("TVD", surveyBlock);
  if (md == null || inc == null || azm == null || n == null || e == null || tvd == null) return null;
  return { md, inc, azm, n, e, tvd };
}

function plannedVector(block: unknown): { md: number; inc: number; azm: number; n: number; e: number; tvd: number } | null {
  const row = Array.isArray(block) ? (Array.isArray(block[0]) ? block[0] : block) : null;
  if (!row) return null;
  const md = num(row[0]);
  if (md == null) return null;
  return { md, inc: num(row[1]) ?? 0, azm: num(row[2]) ?? 0, n: num(row[3]) ?? 0, e: num(row[4]) ?? 0, tvd: 0 };
}

function tdField(field: unknown, vector: { md: number; inc: number; azm: number; n: number; e: number; tvd: number }): number | null {
  switch (String(field ?? "").trim().toUpperCase()) {
    case "MD": return vector.md;
    case "INC": return vector.inc;
    case "AZM":
    case "AZI": return vector.azm;
    case "NS":
    case "N": return vector.n;
    case "EW":
    case "E": return vector.e;
    case "TVD": return vector.tvd;
    default: throw new UdfError("VALUE");
  }
}

function isArgError(value: unknown): boolean {
  return value === ARG_ERROR;
}

export const ARG_ERROR = Symbol("udf-arg-error");

export function loadPlanGrid(header: unknown[], rows: unknown[][]): PlanStation[] {
  const names = header.map((cell) => String(cell ?? "").trim().toUpperCase());
  const index = (name: string) => names.findIndex((cell) => cell === name || (name === "AZM" && cell === "AZI"));
  const mdCol = index("MD");
  const incCol = index("INC");
  const azmCol = index("AZM");
  const tvdCol = index("TVD");
  const nsCol = index("NS");
  const ewCol = index("EW");
  if ([mdCol, incCol, azmCol, tvdCol, nsCol, ewCol].some((value) => value < 0)) return [];
  const plan: PlanStation[] = [];
  for (const row of rows) {
    const md = num(row[mdCol]);
    if (md == null) continue;
    const previous = plan[plan.length - 1];
    if (previous && md <= previous.md) continue;
    plan.push({
      md,
      inc: num(row[incCol]) ?? 0,
      azm: num(row[azmCol]) ?? 0,
      tvd: num(row[tvdCol]) ?? 0,
      n: num(row[nsCol]) ?? 0,
      e: num(row[ewCol]) ?? 0,
    });
  }
  return plan;
}
