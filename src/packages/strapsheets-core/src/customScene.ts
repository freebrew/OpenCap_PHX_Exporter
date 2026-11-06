// The Custom tube as drawing primitives. The phone paints them as SVG; the PDF card turns them into
// vector items. The Entry web app draws the same picture from the same rules in Entry.html.

import type { VectorItem, VectorOp } from "./cardSpec";
import type { Mark } from "./custom";
import {
  BUILDER_H,
  BUILDER_W,
  FLAG_R,
  LENGTH_FLAG_Y,
  OD_FLAG_Y,
  PANEL_THICK,
  PANEL_W,
  PIN_SHOULDER,
  PIN_TIP,
  PIN_W,
  SPIRAL_HALF,
  SPIRAL_W,
  TUBE_CY,
  bodyHalfAt,
  borePolygon,
  featureHalfAt,
  flagPoint,
  flagRow,
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

/** 45-degree hatch lines across a box, cut off at its left and right sides. */
export function hatchLines(left: number, right: number, top: number, bottom: number, step: number): { x1: number; y1: number; x2: number; y2: number }[] {
  const height = bottom - top;
  const lines: { x1: number; y1: number; x2: number; y2: number }[] = [];
  for (let x0 = left - height; x0 < right; x0 += step) {
    const from = Math.max(0, (left - x0) / height);
    const to = Math.min(1, (right - x0) / height);
    if (to <= from) continue;
    lines.push({ x1: x0 + height * from, y1: bottom - height * from, x2: x0 + height * to, y2: bottom - height * to });
  }
  return lines;
}

function pinThreads(mark: Mark, marks: Mark[]): Prim[] {
  const layout = tubeLayout(marks);
  const leftEnd = mark.x < 0.5;
  const x0 = leftEnd ? layout.bodyX - PIN_W : layout.bodyX + layout.bodyW;
  const prims: Prim[] = [];
  for (let index = 1; index <= 4; index += 1) {
    const t = index / 5;
    const px = x0 + PIN_W * t;
    const half = PIN_TIP + (PIN_SHOULDER - PIN_TIP) * (leftEnd ? t : 1 - t);
    prims.push({ t: "line", x1: px, y1: TUBE_CY - half + 1, x2: px, y2: TUBE_CY + half - 1, stroke: INK, w: 1.15 });
  }
  return prims;
}

function spiral(x: number): Prim[] {
  const left = x - SPIRAL_W / 2;
  const right = x + SPIRAL_W / 2;
  const top = TUBE_CY - SPIRAL_HALF;
  const bottom = TUBE_CY + SPIRAL_HALF;
  const prims: Prim[] = [{ t: "poly", pts: rect(left, top, SPIRAL_W, bottom - top), fill: "#ffffff", stroke: "none", w: 0 }];
  for (const line of hatchLines(left, right, top, bottom, 9)) prims.push({ t: "line", ...line, stroke: HATCH, w: 1.4 });
  prims.push({ t: "poly", pts: rect(left, top, SPIRAL_W, bottom - top), fill: "none", stroke: INK, w: 1.8 });
  return prims;
}

function panels(x: number, half: number, color: string): Prim[] {
  const left = x - PANEL_W / 2;
  return [
    { t: "poly", pts: rect(left, TUBE_CY - half + 1, PANEL_W, PANEL_THICK), fill: "#ffffff", stroke: color, w: 1.6 },
    { t: "poly", pts: rect(left, TUBE_CY + half - PANEL_THICK - 1, PANEL_W, PANEL_THICK), fill: "#ffffff", stroke: color, w: 1.6 },
  ];
}

/** Length: a line across the outer diameter, a gap, then a line up to the letter flag. */
function lengthStation(x: number, mark: Mark, marks: Mark[], color: string): Prim[] {
  const half = featureHalfAt(mark.x, marks);
  const top = stationTop(mark, marks);
  return [
    { t: "line", x1: x, y1: TUBE_CY - half, x2: x, y2: TUBE_CY + half, stroke: color, w: 2.4 },
    { t: "line", x1: x - 7, y1: TUBE_CY - half, x2: x + 7, y2: TUBE_CY - half, stroke: color, w: 2.4 },
    { t: "line", x1: x - 7, y1: TUBE_CY + half, x2: x + 7, y2: TUBE_CY + half, stroke: color, w: 2.4 },
    { t: "line", x1: x, y1: top, x2: x, y2: LENGTH_FLAG_Y + FLAG_R, stroke: color, w: 2.4 },
  ];
}

/** OD: a line from the flag up to an arrow that stops short of the tube or feature it points at. */
function odStation(x: number, mark: Mark, marks: Mark[], color: string): Prim[] {
  const tip = stationBottom(mark, marks);
  return [
    { t: "line", x1: x, y1: tip + 9, x2: x, y2: OD_FLAG_Y - FLAG_R, stroke: color, w: 2.4 },
    { t: "poly", pts: [[x, tip], [x - 5.5, tip + 10], [x + 5.5, tip + 10]], fill: color, stroke: color, w: 1 },
  ];
}

function flag(mark: Mark, marks: Mark[], selected: boolean): Prim[] {
  const point = flagPoint(mark, marks);
  const color = selected ? BLUE : NAVY;
  return [
    { t: "circle", cx: point.x, cy: point.y, r: FLAG_R, fill: selected ? BLUE : "#ffffff", stroke: color, w: 2.4 },
    { t: "text", x: point.x, y: point.y + 6, s: mark.letter, size: 18, fill: selected ? "#ffffff" : color },
  ];
}

/** Every primitive of the tube, bottom layer first: tube, pins, add-ons, stations, then the flags. */
export function customScene(marks: Mark[], selected = ""): Prim[] {
  const prims: Prim[] = [{ t: "poly", pts: tubePolygon(marks), fill: STEEL, stroke: INK, w: 2.5 }];
  const bore = borePolygon(marks);
  if (bore.length > 0) prims.push({ t: "poly", pts: bore, fill: "#ffffff", stroke: "none", w: 0 });
  for (const mark of marks) if (mark.kind === "pin") prims.push(...pinThreads(mark, marks));
  for (const mark of marks) {
    const x = flagPoint(mark, marks).x;
    const color = mark.letter === selected ? BLUE : NAVY;
    if (mark.kind === "spiral") prims.push(...spiral(x));
    else if (mark.kind === "panels") prims.push(...panels(x, bodyHalfAt(mark.x, marks), color));
  }
  for (const mark of marks) {
    const x = flagPoint(mark, marks).x;
    const color = mark.letter === selected ? BLUE : NAVY;
    if (mark.kind === "length") prims.push(...lengthStation(x, mark, marks, color));
    else if (mark.kind === "od") prims.push(...odStation(x, mark, marks, color));
  }
  for (const mark of marks) prims.push(...flag(mark, marks, mark.letter === selected));
  return prims;
}

/** The rows of the picture that hold something, so a card can leave out an empty top or bottom. */
export function sceneBounds(marks: Mark[]): { x: number; y: number; w: number; h: number } {
  const rows = new Set(marks.map((mark) => flagRow(mark.kind)));
  const top = rows.has("top") ? 4 : rows.has("mid") ? 54 : 96;
  const bottom = rows.has("bottom") ? BUILDER_H - 2 : TUBE_CY + SPIRAL_HALF + 12;
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
