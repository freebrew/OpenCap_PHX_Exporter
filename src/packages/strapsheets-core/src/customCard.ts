// The Custom.P card: one built tube, the header fields, and A–L. Three cards fit a page.
// Standard families keep the generated card tables; Custom is drawn from its own payload.

import { customLines, markPoint, parseMarks } from "./custom";
import type { EntryPayload } from "./entry";
import { PageCanvas, type PdfWriter } from "./pdfWriter";

const PAGE_WIDTH = 612;
const PAGE_HEIGHT = 792;
const MARGIN_X = 36;
const MARGIN_Y = 36;
const SLOT = 240;

/** Custom pages after the standard families. An empty list adds nothing. */
export function layoutCustomCards(entries: EntryPayload[], writer: PdfWriter): void {
  if (entries.length === 0) return;
  for (let start = 0; start < entries.length; start += 3) {
    const canvas = new PageCanvas({ left: MARGIN_X, top: MARGIN_Y, scale: 0.72, height: PAGE_HEIGHT });
    entries.slice(start, start + 3).forEach((entry, index) => drawCard(canvas, entry, index * SLOT));
    writer.addPage(canvas);
  }
}

function drawCard(canvas: PageCanvas, entry: EntryPayload, top: number): void {
  canvas.text("CUSTOM", 0, top + 18, 16, true, [0, 0, 0]);
  const marks = parseMarks(entry.cells.marksJson);
  canvas.strokeRect(20, top + 28, 520, 70, 1);
  canvas.fillRect(36, top + 50, 488, 26, [0.75, 0.78, 0.82]);
  canvas.line(36, top + 63, 524, top + 63, 1);
  for (const mark of marks) {
    const point = markPoint(mark, marks);
    const x = 36 + (point.x / 640) * 488;
    canvas.text(mark.letter, x, top + 46, 9, true, [0, 0, 0]);
  }
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
    canvas.text(label ?? "", x, top + 112, 8, false, [0.3, 0.3, 0.3]);
    canvas.text(value ?? "", x, top + 126, 10, true, [0, 0, 0]);
  });
  customLines(entry.cells)
    .filter((line) => line.label !== "Total Length" && line.label !== "ID")
    .forEach((line, index) => {
      canvas.text(line.label, index * 40, top + 156, 8, true, [0, 0, 0]);
      canvas.text(line.value, index * 40, top + 168, 8, false, [0, 0, 0]);
    });
}
