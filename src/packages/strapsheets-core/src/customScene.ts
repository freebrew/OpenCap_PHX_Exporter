// The Custom tube as drawing primitives. The phone paints them as SVG; the PDF card turns them into
// vector items. The Entry web app draws the same picture from the same rules in Entry.html.

import type { VectorItem, VectorOp } from "./cardSpec";
import type { EndConfig, Mark } from "./custom";
import {
  BUILDER_W,
  BULGE,
  FLAG_R,
  LENGTH_KNEE_Y,
  OD_KNEE_Y,
  PANEL_THICK,
  PANEL_W,
  PIN_SHOULDER,
  PIN_TIP,
  PIN_W,
  SHAPE_FLAG_Y,
  SEAM_W,
  SPIRAL_HALF,
  TUBE_CY,
  bodyHalfAt,
  borePolygon,
  bulgePolygon,
  featureHalfAt,
  flagPoint,
  flagRadius,
  flagRow,
  lengthFlagY,
  odFlagY,
  spiralPolygon,
  stemPx,
  stationBottom,
  stationTop,
  tubeLayout,
  tubePolygon,
} from "./customShape";

export const INK = "#243038";
export const NAVY = "#1f3b5b";
export const BLUE = "#1a73e8";
/** Neutral grey for a spiral's hatch: the tool's own drawing carries no colour. */
export const HATCH = "#555555";
/** The tube's fill. The phone paints a steel gradient; the PDF uses this flat tone. */
export const STEEL = "steel";
export const STEEL_FLAT = "#d5e0ec";

export type Point = [number, number];

export type Prim =
  | { t: "poly"; pts: Point[]; fill: string; stroke: string; w: number }
  | { t: "line"; x1: number; y1: number; x2: number; y2: number; stroke: string; w: number }
  | { t: "circle"; cx: number; cy: number; r: number; fill: string; stroke: string; w: number }
  | { t: "text"; x: number; y: number; s: string; size: number; fill: string };

function rect(x: number, y: number, w: number, h: number): Point[] {
  return [[x, y], [x + w, y], [x + w, y + h], [x, y + h]];
}

/** The part of a segment that lies inside a convex polygon, or null when none of it does. */
function clipToConvex(poly: Point[], x1: number, y1: number, x2: number, y2: number): { x1: number; y1: number; x2: number; y2: number } | null {
  let area = 0;
  for (let index = 0; index < poly.length; index += 1) {
    const [ax, ay] = poly[index] as Point;
    const [bx, by] = poly[(index + 1) % poly.length] as Point;
    area += ax * by - bx * ay;
  }
  const side = area >= 0 ? 1 : -1;
  let from = 0;
  let to = 1;
  const dx = x2 - x1;
  const dy = y2 - y1;
  for (let index = 0; index < poly.length; index += 1) {
    const [ax, ay] = poly[index] as Point;
    const [bx, by] = poly[(index + 1) % poly.length] as Point;
    const ex = bx - ax;
    const ey = by - ay;
    const start = side * (ex * (y1 - ay) - ey * (x1 - ax));
    const slope = side * (ex * dy - ey * dx);
    if (Math.abs(slope) < 1e-9) {
      if (start < 0) return null;
      continue;
    }
    const t = -start / slope;
    if (slope > 0) from = Math.max(from, t);
    else to = Math.min(to, t);
    if (from > to) return null;
  }
  return { x1: x1 + dx * from, y1: y1 + dy * from, x2: x1 + dx * to, y2: y1 + dy * to };
}

/** 45-degree hatch lines across a convex outline, cut off where they leave it. */
export function hatchLines(poly: Point[], step: number): { x1: number; y1: number; x2: number; y2: number }[] {
  const xs = poly.map((point) => point[0]);
  const ys = poly.map((point) => point[1]);
  const left = Math.min(...xs);
  const right = Math.max(...xs);
  const top = Math.min(...ys);
  const bottom = Math.max(...ys);
  const height = bottom - top;
  const lines: { x1: number; y1: number; x2: number; y2: number }[] = [];
  for (let x0 = left - height; x0 < right; x0 += step) {
    const piece = clipToConvex(poly, x0, bottom, x0 + height, top);
    if (piece && Math.hypot(piece.x2 - piece.x1, piece.y2 - piece.y1) > 1) lines.push(piece);
  }
  return lines;
}
function pinThreads(mark: Mark, marks: Mark[], ink = INK): Prim[] {
  const layout = tubeLayout(marks);
  const leftEnd = mark.x < 0.5;
  const x0 = leftEnd ? layout.bodyX - PIN_W : layout.bodyX + layout.bodyW;
  const prims: Prim[] = [];
  for (let index = 1; index <= 4; index += 1) {
    const t = index / 5;
    const px = x0 + PIN_W * t;
    const half = PIN_TIP + (PIN_SHOULDER - PIN_TIP) * (leftEnd ? t : 1 - t);
    prims.push({ t: "line", x1: px, y1: TUBE_CY - half + 1, x2: px, y2: TUBE_CY + half - 1, stroke: ink, w: 1.15 });
  }
  return prims;
}

/** A spiral: wider than the tube and tapering to the tube's own diameter at both ends. Plain neutral hatch, no colour. */
function spiral(x: number): Prim[] {
  const outline = spiralPolygon(x);
  const prims: Prim[] = [{ t: "poly", pts: outline, fill: "#ffffff", stroke: "none", w: 0 }];
  for (const line of hatchLines(outline, 18)) prims.push({ t: "line", ...line, stroke: HATCH, w: 1.4 });
  prims.push({ t: "poly", pts: outline, fill: "none", stroke: INK, w: 1.8 });
  return prims;
}
/** Seam section: two seam lines across the tube, the outer diameter unchanged. */
function seam(x: number, half: number, color: string): Prim[] {
  return [x - SEAM_W / 2, x + SEAM_W / 2].map((sx): Prim => ({ t: "line", x1: sx, y1: TUBE_CY - half, x2: sx, y2: TUBE_CY + half, stroke: color, w: 2 }));
}

/** Bit box: the wider box end of a motor with the bit screwed into its toe. See the bit-box photos: a make-up shoulder, a white band, then a gold bit head with blades and cutters. */
export const BIT_GOLD = "#f0dfa6";
function bitBox(x: number): Prim[] {
  const spec = BULGE.bitBox as NonNullable<(typeof BULGE)["bitBox"]>;
  const outline = bulgePolygon(x, spec);
  const toe = x - spec.w / 2;
  const at = (dx: number, dy: number): Point => [toe + dx, TUBE_CY + dy];
  const prims: Prim[] = [];
  prims.push({ t: "line", x1: toe + 12, y1: TUBE_CY - spec.half, x2: toe + 12, y2: TUBE_CY + spec.half, stroke: INK, w: 1.4 });
  for (const band of [61, 67]) prims.push({ t: "line", x1: toe + band, y1: TUBE_CY - spec.half, x2: toe + band, y2: TUBE_CY + spec.half, stroke: INK, w: 1.2 });
  prims.push({ t: "poly", pts: outline, fill: "none", stroke: INK, w: 1.8 });
  const head: Point[] = [at(0, -16), at(-9, -22), at(-21, -21), at(-35, -12), at(-41, -4), at(-42, 4), at(-35, 12), at(-21, 21), at(-9, 22), at(0, 16)];
  prims.push({ t: "poly", pts: head, fill: BIT_GOLD, stroke: INK, w: 1.6 });
  for (const [dx, dy] of [[-8, -17], [-20, -15], [-32, -10], [-8, 17], [-20, 15], [-32, 10]] as const) {
    prims.push({ t: "circle", cx: toe + dx, cy: TUBE_CY + dy, r: 2.6, fill: "#ffffff", stroke: INK, w: 1 });
  }
  for (const dy of [-17, 0, 17]) prims.push({ t: "line", x1: toe - 39, y1: TUBE_CY + dy * 0.2, x2: toe - 6, y2: TUBE_CY + dy, stroke: INK, w: 1.2 });
  return prims;
}
/**
 * Adjustable bent housing, drawn straight (the bend is the set angle, not a kink in the picture):
 * upper and lower housing meeting at the joint, a scribe line running across it, and two staggered
 * ladders of alignment ticks either side, as stamped on the housing and orientation sleeve.
 */
function bentHousing(x: number): Prim[] {
  const spec = BULGE.bentHousing as NonNullable<(typeof BULGE)["bentHousing"]>;
  const outline = bulgePolygon(x, spec);
  const cy = TUBE_CY;
  const prims: Prim[] = [
    { t: "poly", pts: outline, fill: "#ffffff", stroke: INK, w: 1.8 },
    { t: "line", x1: x, y1: cy - spec.half, x2: x, y2: cy + spec.half, stroke: INK, w: 1.6 },
  ];
  for (const dy of [-16, -8, 8, 16]) prims.push({ t: "line", x1: x - 3 - 16, y1: cy + dy, x2: x - 3, y2: cy + dy, stroke: INK, w: 1.2 });
  for (const dy of [-12, -4, 4, 12]) prims.push({ t: "line", x1: x + 3, y1: cy + dy, x2: x + 3 + 16, y2: cy + dy, stroke: INK, w: 1.2 });
  prims.push({ t: "line", x1: x - 38, y1: cy, x2: x + 38, y2: cy, stroke: INK, w: 1.8 });
  return prims;
}
function panels(x: number, half: number, color: string): Prim[] {
  const left = x - PANEL_W / 2;
  return [
    { t: "poly", pts: rect(left, TUBE_CY - half + 1, PANEL_W, PANEL_THICK), fill: "#ffffff", stroke: color, w: 1.6 },
    { t: "poly", pts: rect(left, TUBE_CY + half - PANEL_THICK - 1, PANEL_W, PANEL_THICK), fill: "#ffffff", stroke: color, w: 1.6 },
  ];
}

/**
 * Length: a line across the outer diameter, a gap, then a line up to its knee and a leader on to
 * the letter flag, which may sit to one side so close stations keep their flags apart.
 */
function lengthStation(x: number, mark: Mark, marks: Mark[], color: string, touch: boolean): Prim[] {
  const half = featureHalfAt(mark.x, marks);
  const top = stationTop(mark, marks);
  const flagX = flagPoint(mark, marks, touch).x;
  return [
    { t: "line", x1: x, y1: TUBE_CY - half, x2: x, y2: TUBE_CY + half, stroke: color, w: 2.4 },
    { t: "line", x1: x - 7, y1: TUBE_CY - half, x2: x + 7, y2: TUBE_CY - half, stroke: color, w: 2.4 },
    { t: "line", x1: x - 7, y1: TUBE_CY + half, x2: x + 7, y2: TUBE_CY + half, stroke: color, w: 2.4 },
    { t: "line", x1: x, y1: top, x2: x, y2: LENGTH_KNEE_Y, stroke: color, w: 2.4 },
    { t: "line", x1: x, y1: LENGTH_KNEE_Y, x2: flagX, y2: lengthFlagY(touch) + flagRadius(mark.kind, touch), stroke: color, w: 2.4 },
  ];
}

/** OD: a line from the flag up to an arrow that stops short of the tube or feature it points at. */
function odStation(x: number, mark: Mark, marks: Mark[], color: string, touch: boolean): Prim[] {
  const tip = stationBottom(mark, marks);
  const top = stationTop(mark, marks);
  const flagX = flagPoint(mark, marks, touch).x;
  return [
    { t: "line", x1: x, y1: tip + 9, x2: x, y2: OD_KNEE_Y, stroke: color, w: 2.4 },
    { t: "line", x1: x, y1: OD_KNEE_Y, x2: flagX, y2: odFlagY(touch) - flagRadius(mark.kind, touch), stroke: color, w: 2.4 },
    { t: "poly", pts: [[x, tip], [x - 5.5, tip + 10], [x + 5.5, tip + 10]], fill: color, stroke: color, w: 1 },
    // The same arrow from above, so the measure is pinched from both sides.
    { t: "line", x1: x, y1: top - 10, x2: x, y2: top - 24, stroke: color, w: 2.4 },
    { t: "poly", pts: [[x, top], [x - 5.5, top - 10], [x + 5.5, top - 10]], fill: color, stroke: color, w: 1 },
  ];
}

function flag(mark: Mark, marks: Mark[], selected: boolean, touch: boolean, palette: ScenePalette): Prim[] {
  const point = flagPoint(mark, marks, touch);
  const color = selected ? palette.select : palette.navy;
  const radius = flagRadius(mark.kind, touch);
  const size = radius > FLAG_R ? 35 : 18;
  return [
    { t: "circle", cx: point.x, cy: point.y, r: radius, fill: selected ? palette.select : palette.plate, stroke: color, w: radius > FLAG_R ? 3 : 2.4 },
    { t: "text", x: point.x, y: point.y + size / 3, s: mark.letter, size, fill: selected ? palette.plate : color },
  ];
}

/**
 * Every primitive of the tube, bottom layer first: tube, pins, add-ons, stations, then the flags.
 * `touch` draws the builder's finger-size Length and OD flags; the card and the confirm picture leave it off.
 */
export interface ScenePalette {
  ink: string;
  bore: string;
  select: string;
  navy: string;
  plate: string;
}

export const LIGHT_PALETTE: ScenePalette = { ink: INK, bore: "#ffffff", select: BLUE, navy: NAVY, plate: "#ffffff" };

function shortenCopy(text: string): string {
  const value = text.trim();
  return value.length > 18 ? `${value.slice(0, 17)}…` : value;
}

function withEndPins(marks: Mark[], ends: EndConfig): Mark[] {
  const shaped = marks.filter((mark) => mark.kind !== "pin");
  if (ends.lower.toUpperCase() === "PIN") shaped.push({ letter: "", kind: "pin", x: 0.04, value: "" });
  if (ends.upper.toUpperCase() === "PIN") shaped.push({ letter: "", kind: "pin", x: 0.96, value: "" });
  return shaped;
}

function boxThreads(left: boolean, marks: Mark[], ink: string): Prim[] {
  const layout = tubeLayout(marks);
  const x0 = left ? layout.bodyX + 6 : layout.bodyX + layout.bodyW - 20;
  const prims: Prim[] = [];
  for (let row = 0; row < 3; row += 1) {
    const y = TUBE_CY - 8 + row * 8;
    prims.push({ t: "line", x1: x0, y1: y, x2: x0 + 6, y2: y, stroke: ink, w: 1.2 });
    prims.push({ t: "line", x1: x0 + 9, y1: y, x2: x0 + 14, y2: y, stroke: ink, w: 1.2 });
  }
  return prims;
}

function endCaptions(ends: EndConfig, marks: Mark[], ink: string): Prim[] {
  const layout = tubeLayout(marks);
  const leftEdge = layout.pins.left ? layout.bodyX - PIN_W : layout.bodyX;
  const rightEdge = layout.pins.right ? layout.bodyX + layout.bodyW + PIN_W : layout.bodyX + layout.bodyW;
  const leftX = leftEdge + 28;
  const rightX = rightEdge - 28;
  const line = (x: number, y: number, text: string): Prim => ({ t: "text", x, y, s: text, size: 12, fill: ink });
  return [
    line(leftX, TUBE_CY - 66, ends.lower.toUpperCase() === "PIN" ? "PIN" : "BOX"),
    line(leftX, TUBE_CY - 52, shortenCopy(ends.threadLower ?? "")),
    line(rightX, TUBE_CY - 66, ends.upper.toUpperCase() === "PIN" ? "PIN" : "BOX"),
    line(rightX, TUBE_CY - 52, shortenCopy(ends.threadUpper ?? "")),
  ].filter((prim) => prim.t === "text" && prim.s !== "");
}

export function customScene(marks: Mark[], selected = "", touch = false, ends?: EndConfig, palette: ScenePalette = LIGHT_PALETTE): Prim[] {
  const drawn = ends ? marks.filter((mark) => mark.kind !== "pin") : marks;
  const shaped = ends ? withEndPins(drawn, ends) : marks;
  const prims: Prim[] = [{ t: "poly", pts: tubePolygon(shaped, ends), fill: STEEL, stroke: palette.ink, w: 2.5 }];
  const bore = borePolygon(shaped, ends);
  if (bore.length > 0) prims.push({ t: "poly", pts: bore, fill: palette.bore, stroke: "none", w: 0 });
  for (const mark of shaped) if (mark.kind === "pin") prims.push(...pinThreads(mark, shaped, palette.ink));
  if (ends && ends.lower.toUpperCase() !== "PIN") prims.push(...boxThreads(true, shaped, palette.ink));
  if (ends && ends.upper.toUpperCase() !== "PIN") prims.push(...boxThreads(false, shaped, palette.ink));
  if (ends) prims.push(...endCaptions(ends, shaped, palette.ink));
  for (const mark of drawn) {
    const x = stemPx(mark, shaped);
    const color = mark.letter === selected ? palette.select : palette.navy;
    if (mark.kind === "spiral") prims.push(...spiral(x));
    else if (mark.kind === "panels") prims.push(...panels(x, bodyHalfAt(mark.x, drawn), color));
    else if (mark.kind === "seam") prims.push(...seam(x, bodyHalfAt(mark.x, drawn), color));
    else if (mark.kind === "bitBox") prims.push(...bitBox(x));
    else if (mark.kind === "bentHousing") prims.push(...bentHousing(x));
  }
  for (const mark of drawn) {
    const x = stemPx(mark, shaped);
    const color = mark.letter === selected ? palette.select : palette.navy;
    if (mark.kind === "length") prims.push(...lengthStation(x, mark, shaped, color, touch));
    else if (mark.kind === "od") prims.push(...odStation(x, mark, shaped, color, touch));
  }
  for (const mark of drawn) prims.push(...flag(mark, shaped, mark.letter === selected, touch, palette));
  return prims;
}

/** The rows of the picture that hold something, so a card can leave out an empty top or bottom. */
export function sceneBounds(marks: Mark[], ends?: EndConfig): { x: number; y: number; w: number; h: number } {
  const rows = new Set(marks.filter((mark) => mark.kind !== "pin").map((mark) => flagRow(mark.kind)));
  const plain = rows.has("top") ? lengthFlagY() - FLAG_R - 4 : rows.has("mid") ? SHAPE_FLAG_Y - FLAG_R - 4 : TUBE_CY - SPIRAL_HALF - 8;
  const top = ends ? Math.min(plain, TUBE_CY - 80) : plain;
  const bottom = rows.has("bottom") ? odFlagY() + FLAG_R + 2 : TUBE_CY + SPIRAL_HALF + 12;
  return { x: 0, y: top, w: BUILDER_W, h: bottom - top };
}
function circleOps(cx: number, cy: number, r: number): VectorOp[] {
  const k = 0.5523 * r;
  return [
    { op: "M", x: cx + r, y: cy },
    { op: "C", x1: cx + r, y1: cy + k, x2: cx + k, y2: cy + r, x: cx, y: cy + r },
    { op: "C", x1: cx - k, y1: cy + r, x2: cx - r, y2: cy + k, x: cx - r, y: cy },
    { op: "C", x1: cx - r, y1: cy - k, x2: cx - k, y2: cy - r, x: cx, y: cy - r },
    { op: "C", x1: cx + k, y1: cy - r, x2: cx + r, y2: cy - k, x: cx + r, y: cy },
    { op: "Z" },
  ];
}

function flatFill(fill: string): string | undefined {
  if (fill === "none") return undefined;
  return fill === STEEL ? STEEL_FLAT : fill;
}

/** Vector items for the PDF writer, in builder pixels. Text primitives are returned apart. */
export function sceneItems(prims: Prim[]): { items: VectorItem[]; texts: Extract<Prim, { t: "text" }>[] } {
  const items: VectorItem[] = [];
  const texts: Extract<Prim, { t: "text" }>[] = [];
  for (const prim of prims) {
    if (prim.t === "text") {
      texts.push(prim);
    } else if (prim.t === "line") {
      items.push({ kind: "line", x1: prim.x1, y1: prim.y1, x2: prim.x2, y2: prim.y2, stroke: prim.stroke, width: prim.w });
    } else if (prim.t === "circle") {
      items.push({ kind: "path", ops: circleOps(prim.cx, prim.cy, prim.r), fill: flatFill(prim.fill), stroke: prim.stroke === "none" ? undefined : prim.stroke, width: prim.w });
    } else {
      const ops: VectorOp[] = prim.pts.map(([x, y], index) => ({ op: index === 0 ? "M" : "L", x, y }));
      ops.push({ op: "Z" });
      items.push({ kind: "path", ops, fill: flatFill(prim.fill), stroke: prim.stroke === "none" ? undefined : prim.stroke, width: prim.w });
    }
  }
  return { items, texts };
}

