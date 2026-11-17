import { emptyBounds, includeXyz } from "./orbit";
import type { WellScene } from "./wellScene";

// The box the active well casts its shadow on, the way the daily email picture draws it: two walls
// and a floor, padded around the well, with the well's shadow flat on each. The email picture has one
// fixed view, so it has one back wall and one side wall. This view turns, so every wall is built and
// `wallsFor` says which ones are behind the well from where the camera is.

export interface ShadowBox {
  minX: number;
  maxX: number;
  minY: number;
  maxY: number;
  minZ: number;
  maxZ: number;
  /** One grid step for all three axes, metres, so every cell is square. */
  step: number;
}

export type WallSide = "min" | "max";

export interface WallPick {
  /** The wall square to East, on the side away from the camera. */
  x: WallSide;
  /** The wall square to North, on the side away from the camera. */
  y: WallSide;
  /** The floor shows while the camera is above it. */
  floor: boolean;
}

/** Same padding as the email picture: 12% sideways, 8% in depth, never less than 20 m and 4 m. */
const SIDE_PAD = 0.12;
const SIDE_PAD_MIN_M = 20;
const DEPTH_PAD = 0.08;
const DEPTH_PAD_MIN_M = 4;
const DEPTH_SPAN_MIN_M = 12;
const SIDE_SPAN_MIN_M = 80;
const GRID_TICKS = 6;

/** A 1/2/5 grid step giving about the requested tick count. Same rule as the email picture's `niceStep`. */
export function niceStep(span: number, ticks: number): number {
  if (span <= 0 || ticks < 1) return 1;
  const raw = span / ticks;
  const mag = 10 ** Math.floor(Math.log10(raw));
  const m = raw / mag;
  if (m < 1.5) return mag;
  if (m < 3.5) return 2 * mag;
  if (m < 7.5) return 5 * mag;
  return 10 * mag;
}

function widen(low: number, high: number, least: number): [number, number] {
  if (high - low >= least) return [low, high];
  const mid = (low + high) / 2;
  return [mid - least / 2, mid + least / 2];
}

/** The box round the active well: its plan, its drilled path and its target rings. Null with no plan. */
export function shadowBoxOf(scene: WellScene): ShadowBox | null {
  const bounds = emptyBounds();
  includeXyz(bounds, scene.plan.xyz);
  if (scene.drilled) includeXyz(bounds, scene.drilled.xyz);
  for (const target of scene.targets) includeXyz(bounds, target.outer);
  if (!Number.isFinite(bounds.minX)) return null;

  const padX = Math.max(SIDE_PAD_MIN_M, (bounds.maxX - bounds.minX) * SIDE_PAD);
  const padY = Math.max(SIDE_PAD_MIN_M, (bounds.maxY - bounds.minY) * SIDE_PAD);
  const padZ = Math.max(DEPTH_PAD_MIN_M, (bounds.maxZ - bounds.minZ) * DEPTH_PAD);
  const [minZ, maxZ] = widen(bounds.minZ - padZ, bounds.maxZ + padZ, DEPTH_SPAN_MIN_M);
  const side = Math.max(SIDE_SPAN_MIN_M, maxZ - minZ);
  const [minX, maxX] = widen(bounds.minX - padX, bounds.maxX + padX, side);
  const [minY, maxY] = widen(bounds.minY - padY, bounds.maxY + padY, side);
  const step = Math.max(1, niceStep(Math.max(maxX - minX, maxY - minY, maxZ - minZ), GRID_TICKS));
  return { minX, maxX, minY, maxY, minZ, maxZ, step };
}

/** Which walls sit behind the well as seen from the camera. */
export function wallsFor(box: ShadowBox, eye: { x: number; y: number; z: number }): WallPick {
  return {
    x: eye.x >= (box.minX + box.maxX) / 2 ? "min" : "max",
    y: eye.y >= (box.minY + box.maxY) / 2 ? "min" : "max",
    floor: eye.z > box.minZ,
  };
}

/** Grid values from `low` to `high`, on whole steps from zero, so the lines read as round numbers. */
export function gridValues(low: number, high: number, step: number): number[] {
  if (!(step > 0) || !(high > low)) return [];
  const out: number[] = [];
  let value = step * Math.ceil(low / step - 1e-9);
  // `+ 0` turns a negative zero into 0, so a grid line at the origin never prints as -0.
  for (let guard = 0; value <= high + 1e-9 && guard < 200; guard += 1, value += step) out.push(value + 0);
  return out;
}

/** The well's polyline flattened onto one wall, so its shadow reads the same as the well. */
export function flattenOnto(xyz: ArrayLike<number>, axis: "x" | "y" | "z", at: number): Float32Array {
  const out = new Float32Array(xyz.length);
  const index = axis === "x" ? 0 : axis === "y" ? 1 : 2;
  for (let i = 0; i < xyz.length; i += 3) {
    out[i] = xyz[i]!;
    out[i + 1] = xyz[i + 1]!;
    out[i + 2] = xyz[i + 2]!;
    out[i + index] = at;
  }
  return out;
}
