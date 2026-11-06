// The Custom tube's geometry, in builder pixels. One set of rules draws the tube, snaps the Length
// station, keeps flags apart and answers a tap. The Entry web app (Entry.html) and the sheet's
// print picture (custom_diagram.js) follow the same rules in their own units.

import type { Mark, MarkKind } from "./custom";

export const BUILDER_W = 640;
export const BUILDER_H = 260;
export const TUBE_X = 28;
export const TUBE_W = 584;
export const TUBE_CY = 128;
export const BODY_HALF = 16;
export const UPSET_HALF = 10;
/** A spiral (stabilizer or turbolizer) is wider than the tube it sits on. */
export const SPIRAL_HALF = 21;
export const SPIRAL_W = 78;
export const PANEL_W = 58;
export const PANEL_THICK = 6;
export const PIN_W = 28;
export const PIN_TIP = 8;
export const PIN_SHOULDER = 12;

/** Letter flags. A flag is a circle; two flags in one row stay at least FLAG_GAP apart. */
export const FLAG_R = 16;
export const FLAG_GAP = 44;
/** How far a Length stem keeps from the centre of a shape flag it passes. */
export const STEM_CLEAR = 20;
/** A finger lands on a flag inside this radius. */
export const HIT_R = 26;
export const LENGTH_FLAG_Y = 24;
export const SHAPE_FLAG_Y = 76;
export const OD_FLAG_Y = 236;
/** Space left between a station line and the tube or feature it measures. */
export const STATION_GAP = 6;

export interface TubeLayout {
  pins: { left: boolean; right: boolean };
  bodyX: number;
  bodyW: number;
}

export function pinEnds(marks: Mark[]): { left: boolean; right: boolean } {
  let left = false;
  let right = false;
  for (const mark of marks) {
    if (mark.kind !== "pin") continue;
    if (mark.x < 0.5) left = true;
    else right = true;
  }
  return { left, right };
}

export function tubeLayout(marks: Mark[]): TubeLayout {
  const pins = pinEnds(marks);
  const bodyX = TUBE_X + (pins.left ? PIN_W : 0);
  let bodyW = TUBE_W - (pins.left ? PIN_W : 0) - (pins.right ? PIN_W : 0);
  if (bodyW < 80) bodyW = 80;
  return { pins, bodyX, bodyW };
}

function localProfile(xn: number, mx: number, peak: number, span: number, taper: number): number {
  const half = span / 2;
  const left = mx - half;
  const right = mx + half;
  if (xn < left || xn > right) return BODY_HALF;
  if (xn >= left + taper && xn <= right - taper) return peak;
  if (xn < left + taper) return BODY_HALF + (peak - BODY_HALF) * ((xn - left) / taper);
  return peak + (BODY_HALF - peak) * ((xn - (right - taper)) / taper);
}

function splineSpan(mx: number): { left: number; flatRight: number; right: number } {
  const flat = 0.08;
  const taper = 0.05;
  const left = mx - flat * 0.55;
  const flatRight = left + flat;
  return { left, flatRight, right: flatRight + taper };
}

function splineHalf(xn: number, mx: number): number {
  const span = splineSpan(mx);
  if (xn < span.left || xn > span.right) return BODY_HALF;
  if (xn <= span.flatRight) return UPSET_HALF;
  return UPSET_HALF + (BODY_HALF - UPSET_HALF) * ((xn - span.flatRight) / (span.right - span.flatRight));
}

/** Half the tube's height at 0..1 along the body, upsets and splines narrowing it. */
export function bodyHalfAt(xn: number, marks: Mark[]): number {
  let half = BODY_HALF;
  for (const mark of marks) {
    if (mark.kind === "upset") half = Math.min(half, localProfile(xn, mark.x, UPSET_HALF, 0.11, 0.03));
    if (mark.kind === "spline") half = Math.min(half, splineHalf(xn, mark.x));
  }
  return half;
}

/** Where an add-on starts and ends along the body (0..1), or null for a mark with no width. */
export function markExtent(mark: Mark, layout: TubeLayout): [number, number] | null {
  if (mark.kind === "upset") return [mark.x - 0.055, mark.x + 0.055];
  if (mark.kind === "spline") {
    const span = splineSpan(mark.x);
    return [span.left, span.right];
  }
  if (mark.kind === "spiral") return [mark.x - SPIRAL_W / 2 / layout.bodyW, mark.x + SPIRAL_W / 2 / layout.bodyW];
  if (mark.kind === "panels") return [mark.x - PANEL_W / 2 / layout.bodyW, mark.x + PANEL_W / 2 / layout.bodyW];
  return null;
}

/** The outer radius at a point: the tube, or the spiral when one covers the point. */
export function featureHalfAt(xn: number, marks: Mark[]): number {
  const layout = tubeLayout(marks);
  let half = bodyHalfAt(xn, marks);
  for (const mark of marks) {
    if (mark.kind !== "spiral") continue;
    const extent = markExtent(mark, layout);
    if (extent && xn >= extent[0] - 1e-6 && xn <= extent[1] + 1e-6) half = Math.max(half, SPIRAL_HALF);
  }
  return half;
}

function pinHalf(t: number, endHalf: number): number {
  const clamped = Math.min(1, Math.max(0, t));
  const root = Math.min(PIN_SHOULDER, endHalf);
  return PIN_TIP + (root - PIN_TIP) * clamped;
}

function profileBreaks(marks: Mark[]): number[] {
  const xs: number[] = [];
  for (const mark of marks) {
    if (mark.kind === "upset") xs.push(mark.x - 0.055, mark.x - 0.055 + 0.03, mark.x + 0.055 - 0.03, mark.x + 0.055);
    if (mark.kind === "spline") {
      const span = splineSpan(mark.x);
      xs.push(span.left - 0.001, span.left, span.flatRight, span.right);
    }
  }
  return xs;
}

/** The tube's upper edge, left to right, pins included. */
function upperEdge(marks: Mark[], layout: TubeLayout): [number, number][] {
  const points: [number, number][] = [];
  if (layout.pins.left) {
    for (let index = 0; index <= 8; index += 1) {
      const t = index / 8;
      points.push([layout.bodyX - PIN_W + (PIN_W - 0.01) * t, TUBE_CY - pinHalf(t, bodyHalfAt(0, marks))]);
    }
  }
  const xs: number[] = [];
  for (let index = 0; index <= 80; index += 1) xs.push(index / 80);
  for (const xn of profileBreaks(marks)) if (xn > 0 && xn < 1) xs.push(xn);
  xs.sort((a, b) => a - b);
  let previous = -1;
  for (const xn of xs) {
    if (xn - previous < 0.0004) continue;
    previous = xn;
    points.push([layout.bodyX + layout.bodyW * xn, TUBE_CY - bodyHalfAt(xn, marks)]);
  }
  if (layout.pins.right) {
    const end = layout.bodyX + layout.bodyW;
    for (let index = 0; index <= 8; index += 1) {
      const t = index / 8;
      points.push([end + 0.01 + (PIN_W - 0.01) * t, TUBE_CY - pinHalf(1 - t, bodyHalfAt(1, marks))]);
    }
  }
  return points;
}

/** The tube as one closed outline. */
export function tubePolygon(marks: Mark[]): [number, number][] {
  const top = upperEdge(marks, tubeLayout(marks));
  const bottom = [...top].reverse().map(([x, y]): [number, number] => [x, TUBE_CY + (TUBE_CY - y)]);
  return [...top, ...bottom];
}

function halfAtPx(px: number, marks: Mark[], layout: TubeLayout): number {
  if (layout.pins.left && px < layout.bodyX) return pinHalf((px - (layout.bodyX - PIN_W)) / PIN_W, bodyHalfAt(0, marks));
  if (layout.pins.right && px > layout.bodyX + layout.bodyW) {
    return pinHalf(1 - (px - (layout.bodyX + layout.bodyW)) / PIN_W, bodyHalfAt(1, marks));
  }
  const xn = Math.min(1, Math.max(0, (px - layout.bodyX) / layout.bodyW));
  return bodyHalfAt(xn, marks);
}

/** The white bore drawn inside the tube. */
export function borePolygon(marks: Mark[]): [number, number][] {
  const layout = tubeLayout(marks);
  const left = layout.bodyX + 8;
  const right = layout.bodyX + layout.bodyW - 8;
  if (right <= left) return [];
  const top: [number, number][] = [];
  for (let index = 0; index <= 80; index += 1) {
    const px = left + (right - left) * (index / 80);
    top.push([px, TUBE_CY - Math.max(3, halfAtPx(px, marks, layout) * 0.38)]);
  }
  const bottom = [...top].reverse().map(([x, y]): [number, number] => [x, TUBE_CY + (TUBE_CY - y)]);
  return [...top, ...bottom];
}

/** Pixel x of a position along the body. */
export function bodyPx(xn: number, layout: TubeLayout): number {
  return layout.bodyX + xn * layout.bodyW;
}

export type FlagRow = "top" | "mid" | "bottom";

/** Length flags ride the top row, OD flags the bottom row; every other mark sits in the middle row. */
export function flagRow(kind: MarkKind): FlagRow {
  if (kind === "length") return "top";
  if (kind === "od") return "bottom";
  return "mid";
}

/** Centre of a mark's letter flag. A pin's flag sits over the tip of the pin, clear of the shoulder. */
export function flagPoint(mark: Mark, marks: Mark[]): { x: number; y: number } {
  const layout = tubeLayout(marks);
  const row = flagRow(mark.kind);
  const y = row === "top" ? LENGTH_FLAG_Y : row === "bottom" ? OD_FLAG_Y : SHAPE_FLAG_Y;
  if (mark.kind === "pin") return { x: mark.x < 0.5 ? layout.bodyX - PIN_W : layout.bodyX + layout.bodyW + PIN_W, y };
  return { x: bodyPx(mark.x, layout), y };
}

/** Where a Length station's line may stop: both body ends and the left and right edge of every add-on. */
export function lengthSpots(marks: Mark[]): number[] {
  const layout = tubeLayout(marks);
  const spots = [0, 1];
  for (const mark of marks) {
    const extent = markExtent(mark, layout);
    if (!extent) continue;
    for (const edge of extent) if (edge > 0.001 && edge < 0.999) spots.push(edge);
  }
  spots.sort((a, b) => a - b);
  return spots.filter((spot, index) => index === 0 || spot - (spots[index - 1] ?? -1) > 0.002);
}

/** Whether a flag of this kind can sit at xn without crowding another flag or a Length stem. */
export function spotFree(kind: MarkKind, xn: number, marks: Mark[], skipLetter: string): boolean {
  const layout = tubeLayout(marks);
  const x = bodyPx(xn, layout);
  const row = flagRow(kind);
  for (const other of marks) {
    if (other.letter === skipLetter) continue;
    const distance = Math.abs(flagPoint(other, marks).x - x);
    const otherRow = flagRow(other.kind);
    if (otherRow === row && distance < FLAG_GAP) return false;
    if (row === "top" && otherRow === "mid" && distance < STEM_CLEAR) return false;
    if (row === "mid" && otherRow === "top" && distance < STEM_CLEAR) return false;
  }
  return true;
}

/** The free Length spots, nearest to `near` first. */
export function freeLengthSpots(marks: Mark[], near: number, skipLetter: string): number[] {
  return lengthSpots(marks)
    .filter((spot) => spotFree("length", spot, marks, skipLetter))
    .sort((a, b) => Math.abs(a - near) - Math.abs(b - near));
}

/** The y where a station line stops short of what it measures. */
export function stationTop(mark: Mark, marks: Mark[]): number {
  return TUBE_CY - featureHalfAt(mark.x, marks) - STATION_GAP;
}

export function stationBottom(mark: Mark, marks: Mark[]): number {
  return TUBE_CY + featureHalfAt(mark.x, marks) + STATION_GAP;
}

/**
 * The mark under a pointer, in builder pixels, or "". A flag wins over a stem, a stem over the
 * tube behind it; the nearest wins within each.
 */
export function pickMark(px: number, py: number, marks: Mark[]): string {
  let best = "";
  let bestScore = Infinity;
  for (const mark of marks) {
    const flag = flagPoint(mark, marks);
    const flagDistance = Math.hypot(px - flag.x, py - flag.y);
    let score = flagDistance <= HIT_R ? flagDistance : Infinity;
    const across = Math.abs(px - flag.x);
    if (mark.kind === "length" && across <= 14 && py >= flag.y && py <= TUBE_CY + featureHalfAt(mark.x, marks)) score = Math.min(score, HIT_R + across);
    if (mark.kind === "od" && across <= 14 && py >= stationBottom(mark, marks) - STATION_GAP && py <= flag.y) score = Math.min(score, HIT_R + across);
    const body = Math.hypot(px - flag.x, py - TUBE_CY);
    if (mark.kind !== "length" && mark.kind !== "od" && across <= 24 && Math.abs(py - TUBE_CY) <= 30) score = Math.min(score, 2 * HIT_R + body);
    if (score < bestScore) {
      bestScore = score;
      best = mark.letter;
    }
  }
  return best;
}
