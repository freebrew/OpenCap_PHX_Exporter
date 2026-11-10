// Custom tubular marks. The rules match the Entry web app: twelve letters, one pin at each
// end, pins snapped to that end, and every other mark kept on the tube.

import {
  BUILDER_H,
  bitBoxX,
  BUILDER_W,
  flagPoint,
  freeLengthSpots,
  pinEnds,
  spotFree,
  tubeLayout,
  tubePolygon,
} from "./customShape";
import type { Cells } from "./wizard";

export { pinEnds };

export const CUSTOM_LETTERS = "ABCDEFGHIJKL";
export const MARK_KINDS = ["pin", "upset", "spline", "spiral", "panels", "seam", "bentHousing", "bitBox", "length", "od"] as const;
export type MarkKind = (typeof MARK_KINDS)[number];
/** Kinds the Add menu offers. A pin is legacy data only; BOX PIN Config draws the ends. */
export const MENU_KINDS = ["upset", "spline", "spiral", "panels", "seam", "bentHousing", "bitBox", "length", "od"] as const;

export interface EndConfig {
  lower: string;
  upper: string;
  threadLower?: string;
  threadUpper?: string;
}

export interface Mark {
  letter: string;
  kind: MarkKind;
  /** 0 at the upper end, 1 at the lower end. */
  x: number;
  value: string;
  /** A Bent Housing's bend setting in degrees, one of BEND_SETTINGS. */
  setting?: string;
  /** A Bent Housing's type: adjustable (ADJ) or fixed (FXD). */
  housing?: string;
  /** A Bent Housing's length, measured to its centre. */
  length?: string;
}

export function markKindLabel(kind: string): string {
  if (kind === "od") return "OD";
  if (kind === "pin") return "Pin";
  if (kind === "upset") return "Upset";
  if (kind === "spline") return "Spline";
  if (kind === "spiral") return "Spiral";
  if (kind === "panels") return "Panels";
  if (kind === "seam") return "Seam";
  if (kind === "bentHousing") return "Bent Housing";
  if (kind === "bitBox") return "Bit Box";
  if (kind === "length") return "Length";
  return kind;
}

/** The names in the Add menu, in Proper Case. The chips and the confirm lines use the short `markKindLabel`. */
export function markMenuLabel(kind: MarkKind): string {
  switch (kind) {
    case "pin":
      return "Pin";
    case "upset":
      return "Upset / Inset";
    case "spline":
      return "Spline";
    case "spiral":
      return "Spiral (Stab / Turbolizer)";
    case "panels":
      return "Panels";
    case "seam":
      return "Seam Section";
    case "bentHousing":
      return "Bent Housing";
    case "bitBox":
      return "Bit Box";
    case "length":
      return "Length Station";
    case "od":
      return "OD Station";
    default: {
      const never: never = kind;
      return never;
    }
  }
}

/** The settings an adjustable bent housing's orientation sleeve offers, in degrees. */
export const BEND_SETTINGS: readonly string[] = ["0.00", "0.39", "0.78", "1.15", "1.50", "1.83", "2.12", "2.38", "2.60", "2.77", "2.89", "3.00"];

/** Adjustable or fixed bent housing. */
export const BEND_TYPES: readonly string[] = ["ADJ", "FXD"];

/** What an add-on's OD is called: the Bent Housing's is its Pad OD, measured at the joint in the middle. */
export function markOdLabel(kind: MarkKind): string {
  return kind === "bentHousing" ? "Pad OD" : `${markKindLabel(kind)} OD`;
}

/** The add-ons that sit on the tube. The letter they come with is their OD variable. */
export const ADDON_KINDS: readonly MarkKind[] = ["upset", "spline", "spiral", "panels", "seam", "bentHousing", "bitBox"];

export function isAddon(kind: MarkKind): boolean {
  return ADDON_KINDS.includes(kind);
}

/** What a mark's measurement is: a Length station's length, the OD of an OD station or an add-on. */
export function valueLabel(kind: MarkKind): string {
  if (kind === "length") return "Length";
  return "OD";
}

/** Pins stay readable in old JSON and drop out of the working tube. */
export function activeMarks(marks: Mark[]): Mark[] {
  return marks.filter((mark) => mark.kind !== "pin");
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
      const record = item as { letter?: unknown; kind?: unknown; x?: unknown; value?: unknown; setting?: unknown; housing?: unknown; length?: unknown };
      const kind = kindOf(String(record.kind ?? ""));
      const letter = String(record.letter ?? "");
      const x = Number(record.x);
      if (!kind || !/^[A-L]$/.test(letter) || !Number.isFinite(x)) continue;
      const mark: Mark = { letter, kind, x, value: typeof record.value === "string" ? record.value : "" };
      if (kind === "bentHousing" && typeof record.setting === "string") mark.setting = record.setting;
      if (kind === "bentHousing" && typeof record.housing === "string") mark.housing = record.housing;
      if (kind === "bentHousing" && typeof record.length === "string") mark.length = record.length;
      marks.push(mark);
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

export function defaultMarkX(kind: MarkKind, marks: Mark[]): number | null {
  if (kind === "pin") {
    const ends = pinEnds(marks);
    if (!ends.left) return 0.04;
    if (!ends.right) return 0.96;
    return null;
  }
  if (kind === "upset") return 0.62;
  if (kind === "spiral" || kind === "panels" || kind === "seam" || kind === "bentHousing") return 0.5;
  if (kind === "bitBox") return bitBoxX(marks);
  if (kind === "spline") return 0.32;
  const used = marks.filter((mark) => mark.kind !== "pin").length;
  const base = 0.36 + used * 0.12;
  return base > 0.78 ? 0.5 : base;
}

/**
 * Where a drag lands. A pin takes the free end it was dropped toward. A Length station snaps to the
 * nearest free feature of the tube: the edge before the pin threads (or the box end) and the left
 * and right edge of each add-on. Anything else moves freely along the body.
 */
export function snapMarkX(kind: MarkKind, x: number, marks: Mark[], letter: string): number {
  if (kind === "pin") {
    const ends = pinEnds(marks.filter((mark) => mark.letter !== letter));
    if (x < 0.5 && !ends.left) return 0.04;
    if (x >= 0.5 && !ends.right) return 0.96;
    if (!ends.left) return 0.04;
    if (!ends.right) return 0.96;
    return x < 0.5 ? 0.04 : 0.96;
  }
  if (kind === "length") {
    const free = freeLengthSpots(marks, x, letter)[0];
    if (free != null) return free;
    return marks.find((mark) => mark.letter === letter)?.x ?? Math.min(1, Math.max(0, x));
  }
  if (kind === "bitBox") return bitBoxX(marks);
  if (x < 0.04) return 0.04;
  if (x > 0.96) return 0.96;
  return x;
}

/**
 * Letters follow position: the mark furthest left is A, the one furthest right has the highest letter,
 * and those between run in order, whatever order they were made or dragged in. The list keeps its own
 * order, so a mark is found at the same index before and after. A tube with a legacy pin keeps its letters.
 */
export function relabel(marks: Mark[]): Mark[] {
  if (marks.some((mark) => mark.kind === "pin")) return marks;
  const order = marks
    .map((mark, index) => ({ index, x: mark.x, letter: mark.letter }))
    .sort((a, b) => a.x - b.x || a.letter.localeCompare(b.letter) || a.index - b.index);
  const letters = new Map<number, string>();
  order.forEach((item, rank) => letters.set(item.index, CUSTOM_LETTERS[rank] ?? item.letter));
  return marks.map((mark, index) => ({ ...mark, letter: letters.get(index) ?? mark.letter }));
}

/** Kinds a tube holds only one of. */
const SINGLE_KINDS: readonly MarkKind[] = ["bitBox", "bentHousing"];

/** The first spot for a new mark that keeps its flag off every other flag, or null when none is left. */
export function spawnX(kind: MarkKind, marks: Mark[]): number | null {
  const want = defaultMarkX(kind, marks);
  if (want == null || kind === "pin") return want;
  if (kind === "length") return freeLengthSpots(marks, want, "")[0] ?? null;
  for (let step = 0; step <= 48; step += 1) {
    for (const sign of step === 0 ? [1] : [1, -1]) {
      const x = Math.round((want + sign * step * 0.02) * 1000) / 1000;
      if (x >= 0.04 && x <= 0.96 && spotFree(kind, x, marks, "")) return x;
    }
  }
  return null;
}

export function addMark(marks: Mark[], kind: MarkKind): Mark[] | string {
  if (marks.length >= CUSTOM_LETTERS.length) return "A tube holds 12 marks.";
  const letter = nextLetter(marks);
  if (SINGLE_KINDS.includes(kind) && marks.some((mark) => mark.kind === kind)) return `A tube holds one ${markKindLabel(kind)}.`;
  const x = spawnX(kind, marks);
  if (letter === "") return "A tube holds 12 marks.";
  if (x == null) {
    if (kind === "pin") return "Only one pin fits at each end.";
    if (kind === "length") return "No free edge on the tube for another Length station.";
    return `No free spot on the tube for another ${markKindLabel(kind)}.`;
  }
  return relabel([...marks, { letter, kind, x, value: "" }]);
}

export function moveMark(marks: Mark[], letter: string, x: number): Mark[] {
  return relabel(marks.map((mark) => (mark.letter === letter ? { ...mark, x: snapMarkX(mark.kind, x, marks, letter) } : mark)));
}

export function setMarkValue(marks: Mark[], letter: string, value: string): Mark[] {
  return marks.map((mark) => (mark.letter === letter ? { ...mark, value } : mark));
}

export function setMarkSetting(marks: Mark[], letter: string, setting: string): Mark[] {
  return marks.map((mark) => (mark.letter === letter ? { ...mark, setting } : mark));
}

export function setMarkHousing(marks: Mark[], letter: string, housing: string): Mark[] {
  return marks.map((mark) => (mark.letter === letter ? { ...mark, housing } : mark));
}

export function setMarkLength(marks: Mark[], letter: string, length: string): Mark[] {
  return marks.map((mark) => (mark.letter === letter ? { ...mark, length } : mark));
}

export function removeMark(marks: Mark[], letter: string): Mark[] {
  return relabel(marks.filter((mark) => mark.letter !== letter));
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
  for (const mark of activeMarks(parseMarks(cells.marksJson))) {
    if (signedDecimal(mark.value) == null) problems.push({ stepId: "builder", message: `${mark.letter} needs a number.` });
    if (mark.kind === "bentHousing" && !BEND_TYPES.includes(mark.housing ?? "")) problems.push({ stepId: "builder", message: `${mark.letter} needs ADJ or FXD.` });
    if (mark.kind === "bentHousing" && !BEND_SETTINGS.includes(mark.setting ?? "")) problems.push({ stepId: "builder", message: `${mark.letter} needs a Bent Setting.` });
    if (mark.kind === "bentHousing" && signedDecimal(mark.length ?? "") == null) problems.push({ stepId: "builder", message: `${mark.letter} Length needs a number.` });
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
  const marks = activeMarks(parseMarks(cells.marksJson)).map((mark) => {
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
  for (const mark of [...activeMarks(parseMarks(cells.marksJson))].sort((a, b) => a.letter.localeCompare(b.letter))) {
    const name = isAddon(mark.kind) ? markOdLabel(mark.kind) : markKindLabel(mark.kind);
    lines.push({ label: `${mark.letter} ${name}`, value: mark.value });
    if (mark.kind === "bentHousing" && mark.length) lines.push({ label: `${mark.letter} Length To Center`, value: mark.length });
    if (mark.kind === "bentHousing" && mark.housing) lines.push({ label: `${mark.letter} Bent Type`, value: mark.housing });
    if (mark.kind === "bentHousing" && mark.setting) lines.push({ label: `${mark.letter} Bent Setting`, value: mark.setting });
  }
  return lines;
}

/**
 * What the Custom card prints under the header, like the sheet's second row: one cell per lettered
 * flag, in its A–L column. A Length station reads "G LEN"; an OD station and each add-on (whose letter
 * is its OD variable) read "I OD". Pins carry no measurement.
 */
export function customFlagCells(marks: Mark[]): { slot: number; label: string; value: string }[] {
  const cells: { slot: number; label: string; value: string }[] = [];
  for (const mark of marks) {
    if (mark.kind === "pin") continue;
    const slot = CUSTOM_LETTERS.indexOf(mark.letter);
    if (slot < 0) continue;
    cells.push({ slot, label: `${mark.letter} ${mark.kind === "length" ? "LEN" : "OD"}`, value: mark.value });
    if (mark.kind === "bentHousing" && mark.length) cells.push({ slot, label: `${mark.letter} LEN`, value: mark.length });
    if (mark.kind === "bentHousing" && mark.setting) cells.push({ slot, label: `${mark.letter} ${mark.housing || "BEND"}`, value: mark.setting });
  }
  return cells.sort((a, b) => a.slot - b.slot);
}

/** SVG path of the tube, in builder pixels. */
export function tubePath(marks: Mark[]): string {
  const points = tubePolygon(marks);
  if (points.length === 0) return "";
  return `M ${points.map(([x, y]) => `${x.toFixed(2)} ${y.toFixed(2)}`).join(" L ")} Z`;
}

/** The centre of a mark's letter flag in builder pixels. */
export function markPoint(mark: Mark, marks: Mark[]): { x: number; y: number } {
  return flagPoint(mark, marks);
}

export const BUILDER_SIZE = { w: BUILDER_W, h: BUILDER_H };


