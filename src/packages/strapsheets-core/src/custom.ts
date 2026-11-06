// Custom tubular marks. The rules match the Entry web app: twelve letters, one pin at each
// end, pins snapped to that end, and every other mark kept on the tube.

import type { Cells } from "./wizard";

export const CUSTOM_LETTERS = "ABCDEFGHIJKL";
export const MARK_KINDS = ["pin", "upset", "spline", "spiral", "panels", "length", "od"] as const;
export type MarkKind = (typeof MARK_KINDS)[number];

export interface Mark {
  letter: string;
  kind: MarkKind;
  /** 0 at the upper end, 1 at the lower end. */
  x: number;
  value: string;
}

const BUILDER_W = 640;
const TUBE_X = 28;
const TUBE_W = 584;
const TUBE_CY = 128;
const BODY_HALF = 16;
const UPSET_HALF = 10;
const PIN_W = 28;
const PIN_TIP = 8;
const PIN_SHOULDER = 12;

export function markKindLabel(kind: string): string {
  if (kind === "od") return "OD";
  if (kind === "pin") return "Pin";
  if (kind === "upset") return "Upset";
  if (kind === "spline") return "Spline";
  if (kind === "spiral") return "Spiral";
  if (kind === "panels") return "Panels";
  if (kind === "length") return "Length";
  return kind;
}

/** A signed decimal, blank, or null when the text is not a number. */
export function signedDecimal(value: string): string | null {
  const trimmed = value.trim();
  if (trimmed === "") return "";
  return /^-?\d*\.?\d+$/.test(trimmed) && trimmed !== "." && trimmed !== "-" && trimmed !== "-." ? trimmed : null;
}

function kindOf(kind: string): MarkKind | null {
  const text = kind === "mandrel" ? "spline" : kind;
  return MARK_KINDS.includes(text as MarkKind) ? (text as MarkKind) : null;
}

/** Marks from saved JSON. A bad blob is an empty tube, never a throw. */
export function parseMarks(raw: string | undefined): Mark[] {
  try {
    const parsed = JSON.parse(raw && raw.trim() !== "" ? raw : "{}") as { marks?: unknown } | unknown[];
    const list = Array.isArray(parsed) ? parsed : parsed && Array.isArray(parsed.marks) ? parsed.marks : [];
    const marks: Mark[] = [];
    for (const item of list) {
      if (!item || typeof item !== "object") continue;
      const record = item as { letter?: unknown; kind?: unknown; x?: unknown; value?: unknown };
      const kind = kindOf(String(record.kind ?? ""));
      const letter = String(record.letter ?? "");
      const x = Number(record.x);
      if (!kind || !/^[A-L]$/.test(letter) || !Number.isFinite(x)) continue;
      marks.push({ letter, kind, x, value: typeof record.value === "string" ? record.value : "" });
    }
    return marks;
  } catch {
    return [];
  }
}

export function marksJson(marks: Mark[]): string {
  return JSON.stringify({ marks });
}

export function nextLetter(marks: Mark[]): string {
  const used = new Set(marks.map((mark) => mark.letter));
  for (const letter of CUSTOM_LETTERS) {
    if (!used.has(letter)) return letter;
  }
  return "";
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

export function defaultMarkX(kind: MarkKind, marks: Mark[]): number | null {
  if (kind === "pin") {
    const ends = pinEnds(marks);
    if (!ends.left) return 0.04;
    if (!ends.right) return 0.96;
    return null;
  }
  if (kind === "upset") return 0.62;
  if (kind === "spiral" || kind === "panels") return 0.5;
  if (kind === "spline") return 0.32;
  const used = marks.filter((mark) => mark.kind !== "pin").length;
  const base = 0.36 + used * 0.12;
  return base > 0.78 ? 0.5 : base;
}

/** Where a drag lands. A pin takes the free end it was dropped toward. */
export function snapMarkX(kind: MarkKind, x: number, marks: Mark[], letter: string): number {
  if (kind === "pin") {
    const ends = pinEnds(marks.filter((mark) => mark.letter !== letter));
    if (x < 0.5 && !ends.left) return 0.04;
    if (x >= 0.5 && !ends.right) return 0.96;
    if (!ends.left) return 0.04;
    if (!ends.right) return 0.96;
    return x < 0.5 ? 0.04 : 0.96;
  }
  if (x < 0.04) return 0.04;
  if (x > 0.96) return 0.96;
  return x;
}

export function addMark(marks: Mark[], kind: MarkKind): Mark[] | string {
  if (marks.length >= CUSTOM_LETTERS.length) return "A tube holds 12 marks.";
  const letter = nextLetter(marks);
  const x = defaultMarkX(kind, marks);
  if (letter === "" || x == null) return "Only one pin fits at each end.";
  return [...marks, { letter, kind, x, value: "" }];
}

export function moveMark(marks: Mark[], letter: string, x: number): Mark[] {
  return marks.map((mark) => (mark.letter === letter ? { ...mark, x: snapMarkX(mark.kind, x, marks, letter) } : mark));
}

export function setMarkValue(marks: Mark[], letter: string, value: string): Mark[] {
  return marks.map((mark) => (mark.letter === letter ? { ...mark, value } : mark));
}

export function removeMark(marks: Mark[], letter: string): Mark[] {
  return marks.filter((mark) => mark.letter !== letter);
}

/** Pointer x on the builder SVG, as a 0–1 position along the tube body. */
export function xFromClient(clientX: number, rectLeft: number, rectWidth: number, marks: Mark[]): number {
  if (rectWidth <= 0) return 0.5;
  const layout = tubeLayout(marks);
  const px = ((clientX - rectLeft) / rectWidth) * BUILDER_W;
  if (layout.bodyW <= 0) return 0.5;
  return (px - layout.bodyX) / layout.bodyW;
}

export function customProblems(cells: Cells): { stepId: string; message: string }[] {
  const problems: { stepId: string; message: string }[] = [];
  const length = signedDecimal(cells.totalLength ?? "");
  if (length == null) problems.push({ stepId: "builder", message: "Total Length needs a number." });
  else if (length === "") problems.push({ stepId: "builder", message: "Total Length is required." });
  const id = signedDecimal(cells.internalId ?? "");
  if (id == null) problems.push({ stepId: "builder", message: "ID needs a number." });
  for (const mark of parseMarks(cells.marksJson)) {
    if (signedDecimal(mark.value) == null) problems.push({ stepId: "builder", message: `${mark.letter} needs a number.` });
  }
  return problems;
}

/** Named Custom cells ready to store. Blank measurements are dropped; bad text is kept so it can be reported. */
export function normalizeCustomCells(cells: Cells): Cells {
  const next: Cells = {};
  const length = signedDecimal(cells.totalLength ?? "");
  if (length) next.totalLength = length;
  else if (length == null && cells.totalLength) next.totalLength = cells.totalLength;
  const id = signedDecimal(cells.internalId ?? "");
  if (id) next.internalId = id;
  else if (id == null && cells.internalId) next.internalId = cells.internalId;
  const marks = parseMarks(cells.marksJson).map((mark) => {
    const value = signedDecimal(mark.value);
    return value == null ? mark : { ...mark, value };
  });
  if (marks.length > 0 || cells.marksJson) next.marksJson = marksJson(marks);
  return next;
}

export function customLines(cells: Cells): { label: string; value: string }[] {
  const lines: { label: string; value: string }[] = [];
  if (cells.totalLength) lines.push({ label: "Total Length", value: cells.totalLength });
  if (cells.internalId) lines.push({ label: "ID", value: cells.internalId });
  for (const mark of parseMarks(cells.marksJson)) {
    lines.push({ label: `${mark.letter} ${markKindLabel(mark.kind)}`, value: mark.value });
  }
  return lines;
}

interface TubeLayout {
  pins: { left: boolean; right: boolean };
  bodyX: number;
  bodyW: number;
}

function tubeLayout(marks: Mark[]): TubeLayout {
  const pins = pinEnds(marks);
  const bodyX = TUBE_X + (pins.left ? PIN_W : 0);
  let bodyW = TUBE_W - (pins.left ? PIN_W : 0) - (pins.right ? PIN_W : 0);
  if (bodyW < 80) bodyW = 80;
  return { pins, bodyX, bodyW };
}

function upsetHalf(xn: number, mx: number): number {
  const half = 0.055;
  const taper = 0.03;
  const left = mx - half;
  const right = mx + half;
  if (xn < left || xn > right) return BODY_HALF;
  if (xn >= left + taper && xn <= right - taper) return UPSET_HALF;
  if (xn < left + taper) return BODY_HALF + (UPSET_HALF - BODY_HALF) * ((xn - left) / taper);
  return UPSET_HALF + (BODY_HALF - UPSET_HALF) * ((xn - (right - taper)) / taper);
}

function bodyHalf(xn: number, marks: Mark[]): number {
  let height = BODY_HALF;
  for (const mark of marks) {
    if (mark.kind === "upset" || mark.kind === "spline") height = Math.min(height, upsetHalf(xn, mark.x));
  }
  return height;
}

/** SVG path of the tube, in builder pixels. */
export function tubePath(marks: Mark[]): string {
  const layout = tubeLayout(marks);
  const top: string[] = [];
  const bottom: string[] = [];
  const steps = 48;
  for (let index = 0; index <= steps; index += 1) {
    const xn = index / steps;
    const px = layout.bodyX + xn * layout.bodyW;
    const half = bodyHalf(xn, marks);
    top.push(`${px.toFixed(1)} ${(TUBE_CY - half).toFixed(1)}`);
    bottom.push(`${px.toFixed(1)} ${(TUBE_CY + half).toFixed(1)}`);
  }
  if (layout.pins.left) {
    top.unshift(`${(layout.bodyX - PIN_W).toFixed(1)} ${(TUBE_CY - PIN_TIP).toFixed(1)}`);
    bottom.unshift(`${(layout.bodyX - PIN_W).toFixed(1)} ${(TUBE_CY + PIN_TIP).toFixed(1)}`);
  }
  if (layout.pins.right) {
    const end = layout.bodyX + layout.bodyW + PIN_W;
    top.push(`${end.toFixed(1)} ${(TUBE_CY - PIN_TIP).toFixed(1)}`);
    bottom.push(`${end.toFixed(1)} ${(TUBE_CY + PIN_TIP).toFixed(1)}`);
  }
  return `M ${top.join(" L ")} L ${bottom.reverse().join(" L ")} Z`;
}

/** A mark's badge position in builder pixels. */
export function markPoint(mark: Mark, marks: Mark[]): { x: number; y: number } {
  const layout = tubeLayout(marks);
  const x = mark.kind === "pin"
    ? mark.x < 0.5 ? layout.bodyX - PIN_W * 0.55 : layout.bodyX + layout.bodyW + PIN_W * 0.55
    : layout.bodyX + mark.x * layout.bodyW;
  return { x, y: TUBE_CY - BODY_HALF - 28 };
}

export const BUILDER_SIZE = { w: BUILDER_W, h: 220 };
