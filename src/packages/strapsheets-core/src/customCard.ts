// The Custom.P card: one built tube, the header fields, and A–L, styled with the same
// bordered gray/white header row, yellow heavy-bordered metric row, centered diagram,
// and bottom-border family title as the standard print cards.

import { activeMarks, customFlagCells, parseMarks, type EndConfig } from "./custom";
import { customScene, sceneBounds, sceneItems } from "./customScene";
import { BUILDER_W } from "./customShape";
import type { EntryPayload } from "./entry";
import { textWidth, wrapText } from "./helvetica";
import { PageCanvas, type PdfWriter, type Rgb } from "./pdfWriter";

const PAGE_WIDTH = 612;
const PAGE_HEIGHT = 792;
const MARGIN_X = 50.4;
const MARGIN_Y = 54;
const BOTTOM_BORDER = 18;
const POINTS_PER_PIXEL = 0.75;

const BLACK: Rgb = [0, 0, 0];
const GRAY: Rgb = [0.953, 0.953, 0.953];
const YELLOW: Rgb = [1, 0.949, 0.8];

const LABEL_SIZE = 12;
const VALUE_SIZE = 44 / 3;
const LEN_OD_VALUE_SIZE = 64 / 3;
const TITLE_SIZE = 56;
const MIN_SIZE = 8;
const CELL_PADDING = 3;

const TABLE_WIDTH = 940;
const COLUMNS = 12;
const THIN = 1;
const THICK = 3;
const TITLE_ROW = 64;

const PICTURE_SHARE = 0.82;
const PICTURE_MAX_H = 190;
const PICTURE_GAP = 14;
const ROW_HEIGHTS = [20, 36, 20, 36, 18] as const;

/** Pair 1 merged column groups across the 12-column Custom.P grid: Serial (2), Material (1), TOP (3), BTM (3), LEN (2), ID (1). */
const HEADER_GROUPS: readonly { col: number; span: number; label: string; align: "left" | "center" }[] = [
  { col: 0, span: 2, label: "Serial", align: "left" },
  { col: 2, span: 1, label: "Material", align: "center" },
  { col: 3, span: 3, label: "TOP", align: "center" },
  { col: 6, span: 3, label: "BTM", align: "center" },
  { col: 9, span: 2, label: "LEN", align: "center" },
  { col: 11, span: 1, label: "ID", align: "center" },
];

interface CellArea {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** Custom pages after the standard families. An empty list adds nothing. */
export function layoutCustomCards(entries: EntryPayload[], writer: PdfWriter): void {
  if (entries.length === 0) return;
  const scale = Math.min(POINTS_PER_PIXEL, (PAGE_WIDTH - 2 * MARGIN_X) / TABLE_WIDTH);
  const sheetBottom = (PAGE_HEIGHT - MARGIN_Y - BOTTOM_BORDER) / scale;
  const titleTop = sheetBottom - TITLE_ROW;
  const colW = TABLE_WIDTH / COLUMNS;
  const xs = Array.from({ length: COLUMNS + 1 }, (_, i) => i * colW);

  for (let start = 0; start < entries.length; start += 3) {
    const canvas = new PageCanvas({ left: (PAGE_WIDTH - TABLE_WIDTH * scale) / 2, top: MARGIN_Y, scale, height: PAGE_HEIGHT });
    let y = 0;
    for (const entry of entries.slice(start, start + 3)) {
      y = drawCard(canvas, entry, xs, y);
    }
    drawCustomTitle(canvas, titleTop);
    writer.addPage(canvas);
  }
}

function drawCustomTitle(canvas: PageCanvas, top: number): void {
  const title = "CUSTOM";
  const width = textWidth(title, true, TITLE_SIZE);
  canvas.text(title, (TABLE_WIDTH - width) / 2, top + TITLE_ROW - 2 - 0.21 * TITLE_SIZE, TITLE_SIZE, true, BLACK);
}

function rgbOfHex(hex: string): Rgb {
  const value = Number.parseInt(hex.slice(1), 16);
  return [((value >> 16) & 255) / 255, ((value >> 8) & 255) / 255, (value & 255) / 255];
}

function drawPicture(canvas: PageCanvas, marks: ReturnType<typeof parseMarks>, ends: EndConfig, top: number): number {
  const bounds = sceneBounds(marks, ends);
  const targetW = TABLE_WIDTH * PICTURE_SHARE;
  const scale = Math.min(targetW / BUILDER_W, PICTURE_MAX_H / bounds.h);
  const { items, texts } = sceneItems(customScene(activeMarks(marks), "", false, ends));
  let minX = Infinity;
  let maxX = -Infinity;
  for (const item of items) {
    if (item.kind === "line") {
      minX = Math.min(minX, item.x1, item.x2);
      maxX = Math.max(maxX, item.x1, item.x2);
      continue;
    }
    for (const op of item.ops) {
      if ("x" in op) {
        minX = Math.min(minX, op.x);
        maxX = Math.max(maxX, op.x);
      }
    }
  }
  const visualCenter = Number.isFinite(minX) && Number.isFinite(maxX) ? (minX + maxX) / 2 : BUILDER_W / 2;
  const ox = TABLE_WIDTH / 2 - visualCenter * scale;
  const oy = top - bounds.y * scale;
  canvas.paintItems(items, ox, oy, scale, scale);
  const minEdge = ox + minX * scale;
  const maxEdge = ox + maxX * scale;
  for (const text of texts) {
    const size = text.size * scale;
    const w = textWidth(text.s, true, size);
    const rawX = ox + text.x * scale - w / 2;
    const tx = text.s.length > 1 ? Math.max(minEdge, Math.min(maxEdge - w, rawX)) : rawX;
    canvas.text(text.s, tx, oy + text.y * scale - size * 0.08, size, true, rgbOfHex(text.fill));
  }
  return top + bounds.h * scale;
}

function drawCard(canvas: PageCanvas, entry: EntryPayload, xs: number[], top: number): number {
  const marks = parseMarks(entry.cells.marksJson);
  const ends: EndConfig = {
    lower: entry.cells.boxLower === "PIN" ? "PIN" : "BOX",
    upper: entry.cells.boxUpper === "PIN" ? "PIN" : "BOX",
    threadLower: entry.cells.threadLower ?? "",
    threadUpper: entry.cells.threadUpper ?? "",
  };
  const picBottom = drawPicture(canvas, marks, ends, top);
  const tableTop = picBottom + PICTURE_GAP;
  const rowTop = (row: number) => tableTop + ROW_HEIGHTS.slice(0, row).reduce((sum, h) => sum + h, 0);
  const areaOf = (row: number, col: number, span = 1): CellArea => {
    const left = xs[col] ?? 0;
    const right = xs[col + span] ?? TABLE_WIDTH;
    return { x: left, y: rowTop(row), w: right - left, h: ROW_HEIGHTS[row] ?? 0 };
  };

  // Pair 1: Serial, Material, TOP, BTM, LEN, ID
  const pair1Values = [
    entry.cells.serial ?? "",
    entry.cells.material ?? "",
    [entry.cells.boxUpper, entry.cells.threadUpper].filter(Boolean).join(" "),
    [entry.cells.boxLower, entry.cells.threadLower].filter(Boolean).join(" "),
    entry.cells.totalLength ?? "",
    entry.cells.internalId ?? "",
  ];
  for (const group of HEADER_GROUPS) {
    const labelCell = areaOf(0, group.col, group.span);
    canvas.fillRect(labelCell.x, labelCell.y, labelCell.w, labelCell.h, GRAY);
  }
  const p1Top = rowTop(0);
  const p1Height = ROW_HEIGHTS[0] + ROW_HEIGHTS[1];
  canvas.strokeRect(0, p1Top, TABLE_WIDTH, p1Height, THIN);
  canvas.line(0, p1Top + ROW_HEIGHTS[0], TABLE_WIDTH, p1Top + ROW_HEIGHTS[0], THIN);
  for (const group of HEADER_GROUPS.slice(1)) {
    const x = xs[group.col] ?? 0;
    canvas.line(x, p1Top, x, p1Top + p1Height, THIN);
  }
  HEADER_GROUPS.forEach((group, idx) => {
    drawCell(canvas, areaOf(0, group.col, group.span), group.label, LABEL_SIZE, "center");
    const val = (pair1Values[idx] ?? "").trim();
    const isLegacy = group.label === "Serial" || group.label === "TOP" || group.label === "BTM";
    const size = isLegacy ? VALUE_SIZE : LEN_OD_VALUE_SIZE;
    if (val) drawCell(canvas, areaOf(1, group.col, group.span), val, size, group.align);
  });

  // Pair 2: populated A–L flag columns in a single yellow heavy-bordered metric block (matching standard cards)
  const flags = customFlagCells(marks);
  if (flags.length > 0) {
    const p2Top = rowTop(2);
    const p2Height = ROW_HEIGHTS[2] + ROW_HEIGHTS[3];
    const blockWidth = xs[flags.length] ?? TABLE_WIDTH;
    canvas.fillRect(0, p2Top, blockWidth, p2Height, YELLOW);
    canvas.line(0, p2Top + ROW_HEIGHTS[2], blockWidth, p2Top + ROW_HEIGHTS[2], THIN);
    for (let idx = 1; idx < flags.length; idx++) {
      const x = xs[idx] ?? 0;
      canvas.line(x, p2Top, x, p2Top + p2Height, THIN);
    }
    canvas.strokeRect(0, p2Top, blockWidth, p2Height, THICK);
    flags.forEach((flag, idx) => {
      drawCell(canvas, areaOf(2, idx, 1), flag.label, LABEL_SIZE, "center");
      if (flag.value.trim()) drawCell(canvas, areaOf(3, idx, 1), flag.value.trim(), LEN_OD_VALUE_SIZE, "center");
    });
  }

  return rowTop(4) + ROW_HEIGHTS[4];
}

function drawCell(canvas: PageCanvas, area: CellArea, text: string, size: number, align: "center" | "left"): void {
  const room = area.w - 2 * CELL_PADDING;
  const maxLinesFor = (s: number) => (align === "left" ? 1 : Math.max(1, Math.floor(area.h / (s * 1.2))));
  let fitted = size;
  let lines = wrapText(text, true, fitted, room);
  while (fitted > MIN_SIZE && lines.length > maxLinesFor(fitted)) {
    fitted -= 0.5;
    lines = wrapText(text, true, fitted, room);
  }
  const maxLines = maxLinesFor(fitted);
  if (maxLines === 1 && lines.length > 1) lines = [lines.join("")];
  lines = lines.slice(0, maxLines);
  const lineHeight = fitted * 1.2;
  const blockTop = area.y + (area.h - lines.length * lineHeight) / 2;
  canvas.clipped(area.x, area.y, area.w, area.h, () => {
    lines.forEach((line, index) => {
      const x = align === "left" ? area.x + CELL_PADDING : area.x + (area.w - textWidth(line, true, fitted)) / 2;
      canvas.text(line, x, blockTop + index * lineHeight + lineHeight / 2 + 0.36 * fitted, fitted, true, BLACK);
    });
  });
}
