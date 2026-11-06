// The Custom tube's geometry, in builder pixels. One set of rules draws the tube, snaps the Length
// station, keeps flags apart and answers a tap. The Entry web app (Entry.html) and the sheet's
// print picture (custom_diagram.js) follow the same rules in their own units.

import type { Mark, MarkKind } from "./custom";

export const BUILDER_W = 640;
export const BUILDER_H = 390;
/** Wide enough in from each side that a big Length or OD flag at either end of the tube stays on the canvas. */
export const TUBE_X = 52;
export const TUBE_W = 536;
export const TUBE_CY = 200;
export const BODY_HALF = 16;
export const UPSET_HALF = 10;
/** A spiral (stabilizer or turbolizer) is wider than the tube it sits on, tapering down to the tube at both ends. */
export const SPIRAL_HALF = 24;
/** Toe to toe, tapers included. */
export const SPIRAL_W = 78;
/** How far each end of a spiral ramps from the tube's outer diameter up to the spiral's. */
export const SPIRAL_TAPER = 18;
export const PANEL_W = 58;
export const PANEL_THICK = 6;
export const PIN_W = 28;
export const PIN_TIP = 8;
export const PIN_SHOULDER = 12;

/** Letter flags. A flag is a circle; two flags in one row stay at least FLAG_GAP apart. */
export const FLAG_R = 16;
export const FLAG_GAP = 44;
/**
 * A Length or OD flag is 3x that in the builder so a finger can take hold of it. The picture on the
 * confirm screen, the PDF card and the print sheet keep the compact flag.
 */
export const TOUCH_R = 48;
export const TOUCH_GAP = 100;
/** How far a Length or OD stem keeps from the centre of a shape flag it passes. */
export const STEM_CLEAR = 20;
/** Two Length stations (or two OD stations) never share a spot closer than this. */
export const STEM_GAP = 10;
/** A finger lands on a flag inside this radius. */
export const HIT_R = 26;
export const TOUCH_HIT_R = 56;
/** A Length or OD stem rises or falls to its knee, then a leader runs on to the flag, which may sit off to one side. */
export const LENGTH_KNEE_Y = 128;
export const OD_KNEE_Y = 262;
export const LEADER = 28;
export const SHAPE_FLAG_Y = 152;
/** Where the builder's big Length and OD flags sit. */
export const LENGTH_FLAG_Y = LENGTH_KNEE_Y - LEADER - TOUCH_R;
export const OD_FLAG_Y = OD_KNEE_Y + LEADER + TOUCH_R;
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

/**
 * Half the spiral's height at `offset` pixels from its centre: the tube's own radius at the toes,
 * the spiral's at the shoulders, and the straight ramp between.
 */
export function spiralHalfAt(offset: number): number {
  const reach = SPIRAL_W / 2 - Math.abs(offset);
  if (reach <= 0) return BODY_HALF;
  const rise = Math.min(1, reach / SPIRAL_TAPER);
  return Math.round((BODY_HALF + (SPIRAL_HALF - BODY_HALF) * rise) * 1e6) / 1e6;
}

/** The spiral's outline: toe, shoulder, shoulder, toe along the top, then back along the bottom. */
export function spiralPolygon(x: number): [number, number][] {
  const left = x - SPIRAL_W / 2;
  const right = x + SPIRAL_W / 2;
  return [
    [left, TUBE_CY - BODY_HALF],
    [left + SPIRAL_TAPER, TUBE_CY - SPIRAL_HALF],
    [right - SPIRAL_TAPER, TUBE_CY - SPIRAL_HALF],
    [right, TUBE_CY - BODY_HALF],
    [right, TUBE_CY + BODY_HALF],
    [right - SPIRAL_TAPER, TUBE_CY + SPIRAL_HALF],
    [left + SPIRAL_TAPER, TUBE_CY + SPIRAL_HALF],
    [left, TUBE_CY + BODY_HALF],
  ];
}

/**
 * The outer radius at a point: the tube, or the spiral where one covers the point. A Length line at
 * a spiral's shoulder spans the spiral; at its toe it spans the tube.
 */
export function featureHalfAt(xn: number, marks: Mark[]): number {
  const layout = tubeLayout(marks);
  let half = bodyHalfAt(xn, marks);
  for (const mark of marks) {
    if (mark.kind !== "spiral") continue;
    half = Math.max(half, spiralHalfAt((xn - mark.x) * layout.bodyW));
  }
  return half;
}

/**
 * Where a Length station may stop on one add-on: its toes and shoulders (a panel or spline has no
 * ramp of its own, so just its edges).
 */
export function markSpots(mark: Mark, layout: TubeLayout): number[] {
  if (mark.kind === "upset") return [mark.x - 0.055, mark.x - 0.025, mark.x + 0.025, mark.x + 0.055];
  if (mark.kind === "spline") {
    const span = splineSpan(mark.x);
    return [span.left, span.flatRight, span.right];
  }
  if (mark.kind === "spiral") {
    const toe = SPIRAL_W / 2 / layout.bodyW;
    const shoulder = (SPIRAL_W / 2 - SPIRAL_TAPER) / layout.bodyW;
    return [mark.x - toe, mark.x - shoulder, mark.x + shoulder, mark.x + toe];
  }
  if (mark.kind === "panels") return [mark.x - PANEL_W / 2 / layout.bodyW, mark.x + PANEL_W / 2 / layout.bodyW];
  return [];
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

export type FlagRow = "top" | "mid" | "bottom";

/** Length flags ride the top row, OD flags the bottom row; every other mark sits in the middle row. */
export function flagRow(kind: MarkKind): FlagRow {
  if (kind === "length") return "top";
  if (kind === "od") return "bottom";
  return "mid";
}

/** Radius of a mark's flag. `touch` is the builder, where Length and OD flags are big enough for a finger. */
export function flagRadius(kind: MarkKind, touch = false): number {
  return touch && (kind === "length" || kind === "od") ? TOUCH_R : FLAG_R;
}

/** The finger's reach around a flag. */
export function hitRadius(kind: MarkKind): number {
  return kind === "length" || kind === "od" ? TOUCH_HIT_R : HIT_R;
}

/** Where a Length flag sits, and an OD flag: the knee, a leader, then the flag's own radius. */
export function lengthFlagY(touch = false): number {
  return LENGTH_KNEE_Y - LEADER - flagRadius("length", touch);
}

export function odFlagY(touch = false): number {
  return OD_KNEE_Y + LEADER + flagRadius("od", touch);
}

/** Pixel x of a position along the body. */
export function bodyPx(xn: number, layout: TubeLayout): number {
  return layout.bodyX + xn * layout.bodyW;
}

/** Where a mark's stem meets the tube. A pin's flag sits over the tip of the pin, clear of the shoulder. */
export function stemPx(mark: Mark, marks: Mark[]): number {
  const layout = tubeLayout(marks);
  if (mark.kind === "pin") return mark.x < 0.5 ? layout.bodyX - PIN_W : layout.bodyX + layout.bodyW + PIN_W;
  return bodyPx(mark.x, layout);
}

/**
 * Flag centres for the stems of one row, left to right: each as near its own stem as the gap allows
 * (pool-adjacent-violators on the stems less index * gap), then held on the canvas. A flag may sit
 * off to one side of its stem; the leader line joins them.
 */
export function spreadFlags(stems: number[], gap: number, radius: number): number[] {
  const blocks: { sum: number; size: number }[] = [];
  stems.forEach((stem, index) => {
    blocks.push({ sum: stem - index * gap, size: 1 });
    while (blocks.length > 1) {
      const last = blocks[blocks.length - 1] as { sum: number; size: number };
      const before = blocks[blocks.length - 2] as { sum: number; size: number };
      if (before.sum / before.size <= last.sum / last.size) break;
      blocks.splice(blocks.length - 2, 2, { sum: before.sum + last.sum, size: before.size + last.size });
    }
  });
  const centres: number[] = [];
  for (const block of blocks) {
    for (let step = 0; step < block.size; step += 1) centres.push(block.sum / block.size + centres.length * gap);
  }
  for (let index = 0; index < centres.length; index += 1) {
    centres[index] = Math.max(centres[index] as number, index === 0 ? radius : (centres[index - 1] as number) + gap);
  }
  for (let index = centres.length - 1; index >= 0; index -= 1) {
    centres[index] = Math.min(centres[index] as number, index === centres.length - 1 ? BUILDER_W - radius : (centres[index + 1] as number) - gap);
  }
  return centres;
}

/**
 * Centre of a mark's letter flag. Shape flags sit on their own stem; Length and OD flags in a row
 * keep apart from each other and slide aside from their stems when stations are close together.
 */
export function flagPoint(mark: Mark, marks: Mark[], touch = false): { x: number; y: number } {
  const row = flagRow(mark.kind);
  if (row === "mid") return { x: stemPx(mark, marks), y: SHAPE_FLAG_Y };
  const group = marks
    .filter((other) => flagRow(other.kind) === row)
    .sort((a, b) => stemPx(a, marks) - stemPx(b, marks) || a.letter.localeCompare(b.letter));
  const centres = spreadFlags(group.map((other) => stemPx(other, marks)), touch ? TOUCH_GAP : FLAG_GAP, flagRadius(mark.kind, touch));
  const index = group.findIndex((other) => other.letter === mark.letter);
  return { x: centres[index] ?? stemPx(mark, marks), y: row === "top" ? lengthFlagY(touch) : odFlagY(touch) };
}

/**
 * Where a Length station's line may stop: both ends of the body, and the toes and shoulders of every
 * add-on. The ends of the body are the box end (or the edge before the pin threads) and the far end.
 */
export function lengthSpots(marks: Mark[]): number[] {
  const layout = tubeLayout(marks);
  const spots = [0, 1];
  for (const mark of marks) {
    for (const edge of markSpots(mark, layout)) if (edge > 0.001 && edge < 0.999) spots.push(edge);
  }
  spots.sort((a, b) => a - b);
  return spots.filter((spot, index) => index === 0 || spot - (spots[index - 1] ?? -1) > 0.002);
}

/**
 * Whether a mark of this kind can have its stem at xn without crowding another flag or stem. Shape
 * flags keep FLAG_GAP apart; two Length (or OD) stems keep STEM_GAP apart and the row must still have
 * room for every big flag side by side; a Length stem keeps clear of a shape flag it would pass.
 */
export function spotFree(kind: MarkKind, xn: number, marks: Mark[], skipLetter: string): boolean {
  const x = bodyPx(xn, tubeLayout(marks));
  const row = flagRow(kind);
  let sameRow = 0;
  for (const other of marks) {
    if (other.letter === skipLetter) continue;
    const distance = Math.abs(stemPx(other, marks) - x);
    const otherRow = flagRow(other.kind);
    if (otherRow === row) {
      sameRow += 1;
      if (distance < (row === "mid" ? FLAG_GAP : STEM_GAP)) return false;
    }
    if (row === "top" && otherRow === "mid" && distance < STEM_CLEAR) return false;
    if (row === "mid" && otherRow === "top" && distance < STEM_CLEAR) return false;
  }
  if (row !== "mid" && sameRow * TOUCH_GAP > BUILDER_W - 2 * TOUCH_R) return false;
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
 * The mark under a pointer, in builder pixels, or "". The builder draws the big flags, so a flag is
 * tested at its builder size. A flag wins over a stem, a stem over the tube behind it; the nearest wins within each.
 */
export function pickMark(px: number, py: number, marks: Mark[]): string {
  let best = "";
  let bestScore = Infinity;
  for (const mark of marks) {
    const flag = flagPoint(mark, marks, true);
    const stem = stemPx(mark, marks);
    const reach = hitRadius(mark.kind);
    const flagDistance = Math.hypot(px - flag.x, py - flag.y);
    let score = flagDistance <= reach ? flagDistance : Infinity;
    const across = Math.abs(px - stem);
    if (mark.kind === "length" && across <= 14 && py >= LENGTH_KNEE_Y && py <= TUBE_CY + featureHalfAt(mark.x, marks)) score = Math.min(score, reach + across);
    if (mark.kind === "od" && across <= 14 && py >= stationBottom(mark, marks) - STATION_GAP && py <= OD_KNEE_Y) score = Math.min(score, reach + across);
    const body = Math.hypot(px - stem, py - TUBE_CY);
    if (mark.kind !== "length" && mark.kind !== "od" && across <= 24 && Math.abs(py - TUBE_CY) <= 30) score = Math.min(score, 2 * reach + body);
    if (score < bestScore) {
      bestScore = score;
      best = mark.letter;
    }
  }
  return best;
}

/** A drag moves a mark by exactly as far as the pointer travels, so a flag off to one side of its stem does not jump. */
export function dragX(startX: number, startPx: number, nowPx: number, marks: Mark[]): number {
  const layout = tubeLayout(marks);
  return startX + (nowPx - startPx) / layout.bodyW;
}