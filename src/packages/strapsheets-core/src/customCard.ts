// The Custom.P card: one built tube, the header fields, and A–L. Three cards fit a page.
// Standard families keep the generated card tables; Custom is drawn from its own payload.

import { customLines, parseMarks } from "./custom";
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

function drawPicture(canvas: PageCanvas, marks: ReturnType<typeof parseMarks>, top: number): void {
  const bounds = sceneBounds(marks);
  const scale = Math.min(PICTURE_W / BUILDER_W, PICTURE_MAX_H / bounds.h);
  const ox = 20 + (PICTURE_W - BUILDER_W * scale) / 2;
  const oy = top + PICTURE_TOP - bounds.y * scale;
  const { items, texts } = sceneItems(customScene(marks));
  canvas.paintItems(items, ox, oy, scale, scale);
  for (const text of texts) {
    const size = text.size * scale;
    canvas.text(text.s, ox + text.x * scale - size * 0.36, oy + text.y * scale - size * 0.08, size, true, rgbOfHex(text.fill));
  }
}

function drawCard(canvas: PageCanvas, entry: EntryPayload, top: number): void {
  canvas.text("CUSTOM", 0, top + 18, 16, true, [0, 0, 0]);
  const marks = parseMarks(entry.cells.marksJson);
  drawPicture(canvas, marks, top);
  const base = top + PICTURE_TOP + PICTURE_MAX_H + 14;
  const header = [
    ["Serial", entry.cells.serial ?? ""],
    ["Material", entry.cells.material ?? ""],
    ["TOP", [entry.cells.boxUpper, entry.cells.threadUpper].filter(Boolean).join(" ")],
    ["BTM", [entry.cells.boxLower, entry.cells.threadLower].filter(Boolean).join(" ")],
    ["LEN", entry.cells.totalLength ?? ""],
    ["ID", entry.cells.internalId ?? ""],
  ];
  header.forEach(([label, value], index) => {
    const x = index * 88;
    canvas.text(label ?? "", x, base, 8, false, [0.3, 0.3, 0.3]);
    canvas.text(value ?? "", x, base + 14, 10, true, [0, 0, 0]);
  });
  customLines(entry.cells)
    .filter((line) => line.label !== "Total Length" && line.label !== "ID")
    .forEach((line, index) => {
      canvas.text(line.label, index * 40, base + 44, 8, true, [0, 0, 0]);
      canvas.text(line.value, index * 40, base + 56, 8, false, [0, 0, 0]);
    });
}
