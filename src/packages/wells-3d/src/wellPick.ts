import type { PathArrays } from "./wellScene";

export interface PickPoint {
  x: number;
  y: number;
  z: number;
  md: number;
  tvd: number;
  c2c: number | null;
  sf: number | null;
}

export interface PickPath {
  kind: "plan" | "offset" | "drilled";
  name: string;
  points: PickPoint[];
}

export interface ScreenPoint {
  x: number;
  y: number;
  behind: boolean;
}

export interface HoverPick {
  kind: "plan" | "offset" | "drilled";
  name: string;
  refMd: number;
  tvd: number;
  /** Printed centre-to-centre at the nearest report station. Null on the plan and the drilled path. */
  c2c: number | null;
  /** Printed separation factor. On the plan, the tightest one at the nearest report depth. */
  sf: number | null;
  /** 3D metres from the point under the cursor to the closest point on the plan. */
  measuredM: number;
}

export function unpackPath(path: PathArrays): PickPoint[] {
  const out: PickPoint[] = [];
  for (let i = 0; i < path.md.length; i += 1) {
    const c2c = path.c2c[i]!;
    const sf = path.sf[i]!;
    out.push({
      x: path.xyz[i * 3]!,
      y: path.xyz[i * 3 + 1]!,
      z: path.xyz[i * 3 + 2]!,
      md: path.md[i]!,
      tvd: path.tvd[i]!,
      c2c: Number.isFinite(c2c) ? c2c : null,
      sf: Number.isFinite(sf) ? sf : null,
    });
  }
  return out;
}

function clamp01(value: number): number {
  if (value < 0) return 0;
  if (value > 1) return 1;
  return value;
}

/** Closest point on a polyline, in 3D metres. */
export function closestOnPath(points: PickPoint[], x: number, y: number, z: number): { distance: number; md: number; tvd: number } | null {
  if (points.length === 0) return null;
  const first = points[0]!;
  let best = Math.hypot(x - first.x, y - first.y, z - first.z);
  let md = first.md;
  let tvd = first.tvd;
  for (let i = 0; i < points.length - 1; i += 1) {
    const a = points[i]!;
    const b = points[i + 1]!;
    const abx = b.x - a.x;
    const aby = b.y - a.y;
    const abz = b.z - a.z;
    const len2 = abx * abx + aby * aby + abz * abz;
    const t = len2 === 0 ? 0 : clamp01(((x - a.x) * abx + (y - a.y) * aby + (z - a.z) * abz) / len2);
    const px = a.x + abx * t;
    const py = a.y + aby * t;
    const pz = a.z + abz * t;
    const distance = Math.hypot(x - px, y - py, z - pz);
    if (distance < best) {
      best = distance;
      md = a.md + (b.md - a.md) * t;
      tvd = a.tvd + (b.tvd - a.tvd) * t;
    }
  }
  return { distance: best, md, tvd };
}

function pixelSegment(px: number, py: number, ax: number, ay: number, bx: number, by: number): { distance: number; t: number } {
  const abx = bx - ax;
  const aby = by - ay;
  const len2 = abx * abx + aby * aby;
  const t = len2 === 0 ? 0 : clamp01(((px - ax) * abx + (py - ay) * aby) / len2);
  const x = ax + abx * t;
  const y = ay + aby * t;
  return { distance: Math.hypot(px - x, py - y), t };
}

function tightestSf(paths: PickPath[], md: number): number | null {
  let nearest = Infinity;
  let depth = md;
  for (const path of paths) {
    if (path.kind !== "offset") continue;
    for (const point of path.points) {
      if (point.sf == null || point.sf <= 0) continue;
      const gap = Math.abs(point.md - md);
      if (gap < nearest) {
        nearest = gap;
        depth = point.md;
      }
    }
  }
  if (nearest === Infinity) return null;
  let lowest: number | null = null;
  for (const path of paths) {
    if (path.kind !== "offset") continue;
    for (const point of path.points) {
      if (point.sf == null || point.sf <= 0) continue;
      if (Math.abs(point.md - depth) > 1) continue;
      if (lowest == null || point.sf < lowest) lowest = point.sf;
    }
  }
  return lowest;
}

/**
 * The path under the pointer. Printed C2C and SF come from the nearest report station.
 * Measured Here is the 3D distance from the point under the pointer to the plan.
 */
export function pickHover(
  paths: PickPath[],
  pointer: { x: number; y: number },
  project: (x: number, y: number, z: number) => ScreenPoint | null,
  maxPx: number,
): HoverPick | null {
  const plan = paths.find((path) => path.kind === "plan")?.points ?? [];
  let bestDist = maxPx;
  let best: { path: PickPath; index: number; t: number } | null = null;
  for (const path of paths) {
    if (path.points.length === 0) continue;
    const screens = path.points.map((point) => project(point.x, point.y, point.z));
    for (let i = 0; i < path.points.length; i += 1) {
      const screen = screens[i];
      if (!screen || screen.behind) continue;
      const pointDist = Math.hypot(pointer.x - screen.x, pointer.y - screen.y);
      if (pointDist <= bestDist) {
        bestDist = pointDist;
        best = { path, index: i, t: 0 };
      }
      const next = screens[i + 1];
      if (!next || next.behind) continue;
      const along = pixelSegment(pointer.x, pointer.y, screen.x, screen.y, next.x, next.y);
      if (along.distance <= bestDist) {
        bestDist = along.distance;
        best = { path, index: i, t: along.t };
      }
    }
  }
  if (!best) return null;
  const chosen = best;
  const from = chosen.path.points[chosen.index]!;
  const to = chosen.path.points[chosen.index + 1] ?? from;
  const t = to === from ? 0 : chosen.t;
  const world = {
    x: from.x + (to.x - from.x) * t,
    y: from.y + (to.y - from.y) * t,
    z: from.z + (to.z - from.z) * t,
  };
  const printed = t < 0.5 ? from : to;
  const measured = closestOnPath(plan, world.x, world.y, world.z);
  const onPlan = chosen.path.kind === "plan";
  return {
    kind: chosen.path.kind,
    name: chosen.path.name,
    refMd: printed.md,
    tvd: printed.tvd,
    c2c: onPlan || chosen.path.kind === "drilled" ? null : printed.c2c,
    sf: onPlan ? tightestSf(paths, printed.md) : (chosen.path.kind === "drilled" ? null : printed.sf),
    measuredM: measured?.distance ?? 0,
  };
}
