import { interpCourseDir, minCurveStep } from "./angles";

export interface PlanStation {
  md: number;
  inc: number;
  azm: number;
  tvd: number;
  n: number;
  e: number;
}

export interface PlanPoint {
  n: number;
  e: number;
  tvd: number;
  azm: number;
  inc: number;
}

export function planAt(md: number, plan: PlanStation[]): PlanPoint | null {
  const count = plan.length;
  if (count < 1) return null;
  const first = plan[0];
  const last = plan[count - 1];
  if (!first || !last) return null;
  if (md <= first.md) return { n: first.n, e: first.e, tvd: first.tvd, azm: first.azm, inc: first.inc };
  if (md >= last.md) return { n: last.n, e: last.e, tvd: last.tvd, azm: last.azm, inc: last.inc };
  let index = 0;
  for (let i = 0; i < count - 1; i += 1) {
    const left = plan[i];
    const right = plan[i + 1];
    if (left && right && md >= left.md && md <= right.md) {
      index = i;
      break;
    }
  }
  const left = plan[index];
  const right = plan[index + 1];
  if (!left || !right) return null;
  const fraction = right.md > left.md ? (md - left.md) / (right.md - left.md) : 0;
  if (fraction <= 0) return { n: left.n, e: left.e, tvd: left.tvd, azm: left.azm, inc: left.inc };
  if (fraction >= 1) return { n: right.n, e: right.e, tvd: right.tvd, azm: right.azm, inc: right.inc };
  const direction = interpCourseDir(left.inc, left.azm, right.inc, right.azm, fraction);
  const step = minCurveStep(left.md, left.inc, left.azm, md, direction.inc, direction.azm);
  return {
    n: left.n + step.n,
    e: left.e + step.e,
    tvd: left.tvd + step.tvd,
    azm: direction.azm,
    inc: direction.inc,
  };
}

function planDistance2(md: number, n: number, e: number, tvd: number, plan: PlanStation[]): number {
  const point = planAt(md, plan);
  if (!point) return Number.POSITIVE_INFINITY;
  return (n - point.n) ** 2 + (e - point.e) ** 2 + (tvd - point.tvd) ** 2;
}

/** Closest plan point within 150 m of md, then a ternary refine. */
export function closestPlanPoint(md: number, n: number, e: number, tvd: number, plan: PlanStation[]): PlanPoint | null {
  if (plan.length < 2) return null;
  const first = plan[0];
  const last = plan[plan.length - 1];
  if (!first || !last) return null;
  let lo = Math.max(first.md, md - 150);
  let hi = Math.min(last.md, md + 150);
  let bestMd = lo;
  let best = -1;
  for (let q = lo; q <= hi + 1e-9; q += 1) {
    const distance = planDistance2(q, n, e, tvd, plan);
    if (best < 0 || distance < best) {
      best = distance;
      bestMd = q;
    }
  }
  let a = Math.max(lo, bestMd - 1);
  let b = Math.min(hi, bestMd + 1);
  for (let k = 0; k < 40; k += 1) {
    const m1 = a + (b - a) / 3;
    const m2 = b - (b - a) / 3;
    if (planDistance2(m1, n, e, tvd, plan) < planDistance2(m2, n, e, tvd, plan)) b = m2;
    else a = m1;
  }
  return planAt((a + b) / 2, plan);
}

export interface GaugeOffset {
  /** Left (-)/Right (+) in the plan's normal plane (East when plan is vertical). */
  side: number;
  /** Up (+)/Down (-) along the plan's high-side axis (North when plan is vertical). */
  vertical: number;
  total: number;
}

/**
 * Wellhead-relative N/E from MD / Inc / Azm (minimum curvature), optionally
 * starting from a non-zero tie-on (startN, startE).
 */
export function integrateLocalNE(md: number[], inc: number[], azm: number[], startN = 0, startE = 0): { ns: number[]; ew: number[] } {
  const ns = md.map(() => startN);
  const ew = md.map(() => startE);
  for (let i = 1; i < md.length; i += 1) {
    const step = minCurveStep(md[i - 1]!, inc[i - 1]!, azm[i - 1]!, md[i]!, inc[i]!, azm[i]!);
    ns[i] = ns[i - 1]! + step.n;
    ew[i] = ew[i - 1]! + step.e;
  }
  return { ns, ew };
}

/**
 * Well Seeker X 3D closest-approach proximity in the plan's orthogonal
 * High-Side / Right normal plane (or North-referenced when the closest plan
 * point is vertical).
 */
export function closestOffset(md: number, n: number, e: number, tvd: number, plan: PlanStation[]): GaugeOffset | null {
  const point = closestPlanPoint(md, n, e, tvd, plan);
  if (!point) return null;
  const dN = n - point.n;
  const dE = e - point.e;
  const dTvdUp = point.tvd - tvd;
  if (point.inc > 0.001) {
    const az = point.azm * Math.PI / 180;
    const inc = point.inc * Math.PI / 180;
    const side = dE * Math.cos(az) - dN * Math.sin(az);
    const along = dN * Math.cos(az) + dE * Math.sin(az);
    const vertical = along * Math.cos(inc) + dTvdUp * Math.sin(inc);
    return { side, vertical, total: Math.hypot(side, vertical) };
  }
  return { side: dE, vertical: dN, total: Math.hypot(dE, dN) };
}
