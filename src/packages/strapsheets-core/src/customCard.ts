// The Custom.P card: one built tube, the header fields, and A–L. Three cards fit a page.
// Standard families keep the generated card tables; Custom is drawn from its own payload.

import { activeMarks, customFlagCells, parseMarks, type EndConfig } from "./custom";
import { customScene, sceneBounds, sceneItems } from "./customScene";
import { BUILDER_W } from "./customShape";
import type { EntryPayload } from "./entry";
import { PageCanvas, type PdfWriter, type Rgb } from "./pdfWriter";

const PAGE_WIDTH = 612;
const PAGE_HEIGHT = 792;
const MARGIN_X = 36;
const MARGIN_Y = 36;
const SLOT = 300;
/** Layout units the picture is scaled to fill. */
const PICTURE_W = 500;
const PICTURE_TOP = 26;
const PICTURE_MAX_H = 190;
/** Width of one A–L flag column under the header. */
const FLAG_PITCH = 41;
/** Width of each header column: Serial, Material, TOP, BTM, LEN, ID. The thread columns hold "PIN 4 1/2 IF [NC 50]". */
const HEADER_WIDTHS = [70, 70, 140, 140, 56, 52];

/** Custom pages after the standard families. An empty list adds nothing. */
export function layoutCustomCards(entries: EntryPayload[], writer: PdfWriter): void {
  if (entries.length === 0) return;
  for (let start = 0; start < entries.length; start += 3) {
    const canvas = new PageCanvas({ left: MARGIN_X, top: MARGIN_Y, scale: 0.72, height: PAGE_HEIGHT });
    entries.slice(start, start + 3).forEach((entry, index) => drawCard(canvas, entry, index * SLOT));
    writer.addPage(canvas);
  }
}

function rgbOfHex(hex: string): Rgb {
  const value = Number.parseInt(hex.slice(1), 16);
  return [((value >> 16) & 255) / 255, ((value >> 8) & 255) / 255, (value & 255) / 255];
}

function drawPicture(canvas: PageCanvas, marks: ReturnType<typeof parseMarks>, ends: EndConfig, top: number): void {
  const bounds = sceneBounds(marks, ends);
  const scale = Math.min(PICTURE_W / BUILDER_W, PICTURE_MAX_H / bounds.h);
  const ox = 20 + (PICTURE_W - BUILDER_W * scale) / 2;
  const oy = top + PICTURE_TOP - bounds.y * scale;
  const { items, texts } = sceneItems(customScene(activeMarks(marks), "", false, ends));
  canvas.paintItems(items, ox, oy, scale, scale);
  for (const text of texts) {
    const size = text.size * scale;
    canvas.text(text.s, ox + text.x * scale - size * 0.36, oy + text.y * scale - size * 0.08, size, true, rgbOfHex(text.fill));
  }
}

function drawCard(canvas: PageCanvas, entry: EntryPayload, top: number): void {
  canvas.text("CUSTOM", 0, top + 18, 16, true, [0, 0, 0]);
  const marks = parseMarks(entry.cells.marksJson);
  const ends: EndConfig = {
    lower: entry.cells.boxLower === "PIN" ? "PIN" : "BOX",
    upper: entry.cells.boxUpper === "PIN" ? "PIN" : "BOX",
    threadLower: entry.cells.threadLower ?? "",
    threadUpper: entry.cells.threadUpper ?? "",
  };
  drawPicture(canvas, marks, ends, top);
  const base = top + PICTURE_TOP + PICTURE_MAX_H + 14;
  const header = [
    ["Serial", entry.cells.serial ?? ""],
    ["Material", entry.cells.material ?? ""],
    ["TOP", [entry.cells.boxUpper, entry.cells.threadUpper].filter(Boolean).join(" ")],
    ["BTM", [entry.cells.boxLower, entry.cells.threadLower].filter(Boolean).join(" ")],
    ["LEN", entry.cells.totalLength ?? ""],
    ["ID", entry.cells.internalId ?? ""],
  ];
  let x = 0;
  header.forEach(([label, value], index) => {
    canvas.text(label ?? "", x, base, 8, false, [0.3, 0.3, 0.3]);
    canvas.text(value ?? "", x, base + 14, 10, true, [0, 0, 0]);
    x += HEADER_WIDTHS[index] ?? 0;
  });
  for (const flag of customFlagCells(marks)) {
    const x = flag.slot * FLAG_PITCH;
    canvas.text(flag.label, x, base + 44, 8, true, [0, 0, 0]);
    canvas.text(flag.value, x, base + 56, 8, false, [0, 0, 0]);
  }
}
