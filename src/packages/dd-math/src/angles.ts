const PI = 3.14159265358979;
const EPS = 0.0000001;

export function deg2rad(degrees: number): number {
  return degrees * PI / 180;
}

export function rad2deg(radians: number): number {
  return radians * 180 / PI;
}

/** VBA Int truncates toward negative infinity. */
export function wrap180(value: number): number {
  let next = value - 360 * Math.floor(value / 360);
  if (next > 180) next -= 360;
  return next;
}

export function wrap360(value: number): number {
  let next = value - 360 * Math.floor(value / 360);
  if (next < 0) next += 360;
  return next;
}

export function doglegAngleDeg(inc1: number, azm1: number, inc2: number, azm2: number): number {
  const r1 = deg2rad(inc1);
  const r2 = deg2rad(inc2);
  const dA = deg2rad(wrap180(azm2 - azm1));
  let cosine = Math.cos(r1) * Math.cos(r2) + Math.sin(r1) * Math.sin(r2) * Math.cos(dA);
  if (cosine > 1) cosine = 1;
  if (cosine < -1) cosine = -1;
  return rad2deg(Math.acos(cosine));
}

/** Excel Atan2(x, y) is Math.atan2(y, x). */
export function toolfaceBetweenDeg(inc1: number, azm1: number, inc2: number, azm2: number): number {
  const r1 = deg2rad(inc1);
  const r2 = deg2rad(inc2);
  const dA = deg2rad(wrap180(azm2 - azm1));
  const x = Math.sin(r2) * Math.cos(r1) * Math.cos(dA) - Math.sin(r1) * Math.cos(r2);
  const y = Math.sin(r2) * Math.sin(dA);
  return rad2deg(Math.atan2(y, x));
}

export function interpCourseDir(inc1: number, azm1: number, inc2: number, azm2: number, fraction: number): { inc: number; azm: number } {
  const n1 = Math.sin(deg2rad(inc1)) * Math.cos(deg2rad(azm1));
  const e1 = Math.sin(deg2rad(inc1)) * Math.sin(deg2rad(azm1));
  const v1 = Math.cos(deg2rad(inc1));
  const n2 = Math.sin(deg2rad(inc2)) * Math.cos(deg2rad(azm2));
  const e2 = Math.sin(deg2rad(inc2)) * Math.sin(deg2rad(azm2));
  const v2 = Math.cos(deg2rad(inc2));
  let cosB = n1 * n2 + e1 * e2 + v1 * v2;
  if (cosB > 1) cosB = 1;
  if (cosB < -1) cosB = -1;
  const beta = Math.acos(cosB);
  let n = n1;
  let e = e1;
  let v = v1;
  if (beta >= 0.0000001) {
    const s = Math.sin(beta);
    n = (Math.sin((1 - fraction) * beta) * n1 + Math.sin(fraction * beta) * n2) / s;
    e = (Math.sin((1 - fraction) * beta) * e1 + Math.sin(fraction * beta) * e2) / s;
    v = (Math.sin((1 - fraction) * beta) * v1 + Math.sin(fraction * beta) * v2) / s;
  }
  const mag = Math.sqrt(n * n + e * e + v * v);
  if (mag > 0) {
    n /= mag;
    e /= mag;
    v /= mag;
  }
  if (v > 1) v = 1;
  if (v < -1) v = -1;
  const inc = Math.acos(v) * 180 / PI;
  if (Math.abs(n) < 1e-10 && Math.abs(e) < 1e-10) return { inc, azm: azm1 };
  let azm = Math.atan2(e, n) * 180 / PI;
  if (azm < 0) azm += 360;
  return { inc, azm };
}

export interface CurveStep {
  tvd: number;
  n: number;
  e: number;
}

export function minCurveStep(md1: number, inc1: number, azm1: number, md2: number, inc2: number, azm2: number): CurveStep {
  const beta = deg2rad(doglegAngleDeg(inc1, azm1, inc2, azm2));
  const h = beta > 0.0000001 ? (md2 - md1) / 2 * (2 / beta * Math.tan(beta / 2)) : (md2 - md1) / 2;
  const r1 = deg2rad(inc1);
  const r2 = deg2rad(inc2);
  const b1 = deg2rad(azm1);
  const b2 = deg2rad(azm2);
  return {
    tvd: h * (Math.cos(r1) + Math.cos(r2)),
    n: h * (Math.sin(r1) * Math.cos(b1) + Math.sin(r2) * Math.cos(b2)),
    e: h * (Math.sin(r1) * Math.sin(b1) + Math.sin(r2) * Math.sin(b2)),
  };
}

export function arcToPoint(incBit: number, azmBit: number, dN: number, dE: number, dV: number): { tf: number; dls: number } | null {
  const ri = deg2rad(incBit);
  const ra = deg2rad(azmBit);
  const hs = dN * Math.cos(ri) * Math.cos(ra) + dE * Math.cos(ri) * Math.sin(ra) - dV * Math.sin(ri);
  const rt = -dN * Math.sin(ra) + dE * Math.cos(ra);
  const along = dN * Math.sin(ri) * Math.cos(ra) + dE * Math.sin(ri) * Math.sin(ra) + dV * Math.cos(ri);
  const chord = Math.sqrt(dN * dN + dE * dE + dV * dV);
  if (chord <= EPS) return null;
  const side = Math.sqrt(hs * hs + rt * rt);
  if (side <= 0.000001) return { tf: 0, dls: 0 };
  const theta = Math.atan2(side, along);
  const arcLen = chord * theta / Math.sin(theta);
  if (arcLen <= EPS) return null;
  return {
    tf: rad2deg(Math.atan2(rt, hs)),
    dls: rad2deg(2 * theta) * 30 / arcLen,
  };
}

export function inverseCourse(inc1: number, azm1: number, n1: number, e1: number, v1: number, inc2: number, azm2: number, n2: number, e2: number, v2: number): number {
  const chord = Math.sqrt((n2 - n1) ** 2 + (e2 - e1) ** 2 + (v2 - v1) ** 2);
  const beta = deg2rad(doglegAngleDeg(inc1, azm1, inc2, azm2));
  if (beta > EPS) return chord * beta / (2 * Math.sin(beta / 2));
  return chord;
}

export { EPS, PI };
