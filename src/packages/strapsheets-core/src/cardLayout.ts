// Lays entries out as the Sheets printout does: one drawing per tool family, then a card for each
// entry (a grey label row over a value row, the lengths and diameters in a yellow block with a
// heavy border), then the family title in large type. Dimensions are the sheet's own pixels
// (print_cards.py); the page scale is "fit to width", as the print engine exports.

import { CARD_ART } from "./art.gen";
import { layoutCustomCards } from "./customCard";
import type { CardArt, PrintCard } from "./cardSpec";
import { PRINT_CARDS } from "./cards.gen";
import type { EntryPayload } from "./entry";
import { textWidth, wrapText } from "./helvetica";
import { ENTRY_FAMILIES } from "./families.gen";
import { MWD_WHICH, mwdCardEntry } from "./mwd";
import { PageCanvas, type PdfWriter, type Rgb } from "./pdfWriter";
import { findFamily } from "./wizard";

const PAGE_WIDTH = 612;
const PAGE_HEIGHT = 792;
const MARGIN_X = 50.4;
const MARGIN_Y = 54;
const BOTTOM_BORDER = 18;
/** Points in a sheet pixel at 100%. */
const POINTS_PER_PIXEL = 0.75;

const BLACK: Rgb = [0, 0, 0];
const GRAY: Rgb = [0.953, 0.953, 0.953];
const YELLOW: Rgb = [1, 0.949, 0.8];

/** Arial 9, 11, 16 and 42 point, in sheet pixels. */
const LABEL_SIZE = 12;
const VALUE_SIZE = 44 / 3;
const LARGE_VALUE_SIZE = 64 / 3;
const TITLE_SIZE = 56;
const MIN_SIZE = 8;
const CELL_PADDING = 3;

const SINGLE_WIDTH = 940;
const DUAL_WIDTH = 956;
const DUAL_GAP = 28;
const SERIAL_COLUMN = 148;
const MIN_COLUMN = 52;

const THIN = 1;
const THICK = 3;
const FOOTER_GAP = 42;
const TITLE_ROW = 64;
const DRAWING_SHARE = 0.82;
const DRAWING_MAX_HEIGHT = 300;
const DRAWING_GAP = 14;

/** A length, outside or inside diameter label: the cells that sit in the yellow block. */
const METRIC_LABEL = /^[A-Z] (LEN|OD|ID)$/;
/** Only Serial (and MWD sub serials) and TOP / BTM thread boxes stay at legacy VALUE_SIZE; all other value boxes use LARGE_VALUE_SIZE. */
const LEGACY_VALUE_LABEL = /^(Serial|Carrier|.* SN|TOP|BTM)$/;

/** The layout a Subs entry with no known sub type falls back to, so the entry is never dropped. */
const SUBS_FALLBACK: PrintCard | null = (() => {
  const first = PRINT_CARDS.find((card) => card.typeCol != null);
  return first ? { ...first, name: "Subs", title: "SUBS", typeValues: null } : null;
})();

interface Cell {
  x: number;
  y: number;
  w: number;
  h: number;
}

interface Group {
  card: PrintCard;
  /** Where the group sorts among the print tabs. */
  order: number;
  entries: EntryPayload[];
}

/** The print tab an entry goes on. Subs split by sub type across the Float, GAP and X-Over tabs. */
export function printCardFor(entry: EntryPayload): { card: PrintCard; order: number } | null {
  const candidates = PRINT_CARDS.filter((card) => card.which === entry.which);
  const first = candidates[0];
  if (!first) return null;
  if (candidates.length === 1 || first.typeCol == null) return { card: first, order: PRINT_CARDS.indexOf(first) };
  const type = (entry.cells[first.typeCol] ?? "").trim().toLowerCase();
  const match = candidates.find((card) => card.typeValues?.some((value) => value.toLowerCase() === type));
  if (match) return { card: match, order: PRINT_CARDS.indexOf(match) };
  return SUBS_FALLBACK ? { card: SUBS_FALLBACK, order: PRINT_CARDS.indexOf(first) + 0.5 } : { card: first, order: PRINT_CARDS.indexOf(first) };
}

/** The cells the card reads: the MWD Carrier's named cells become its sheet row, every other family's are already columns. */
function cardEntry(entry: EntryPayload): EntryPayload {
  const family = entry.which === MWD_WHICH ? findFamily(ENTRY_FAMILIES, entry.which) : null;
  return family ? mwdCardEntry(family, entry) : entry;
}

/** The columns a header cell covers: one, or more where the card merges it (the MWD card's Date, TOP and BTM). */
function spanOf(card: PrintCard, row: number, column: number): number {
  if (row > 1) return 1;
  return card.headerSpans?.find(([first]) => first === column)?.[1] ?? 1;
}

/** Pixel width of every column of a card, as the live tab is sized. */
export function columnWidths(card: PrintCard): number[] {
  const dual = card.positions === 2;
  const gap = dual ? DUAL_GAP : 0;
  const each = Math.max(MIN_COLUMN, Math.floor(((dual ? DUAL_WIDTH : SINGLE_WIDTH) - gap) / card.width));
  const serials = new Set(Array.from({ length: card.positions }, (_, position) => card.serialCell[1] + position * card.posShift));
  return Array.from({ length: card.width }, (_, column) => {
    if (serials.has(column)) return SERIAL_COLUMN;
    if (dual && column === card.posShift - 1) return gap;
    return each;
  });
}

/** Add the pages for every entry: families in the order the sheet prints them, one page break between. */
export function layoutCards(entries: EntryPayload[], writer: PdfWriter): void {
  const groups: Group[] = [];
  const unknown: EntryPayload[] = [];
  const custom = entries.filter((entry) => entry.which === "Custom");
  for (const source of entries) {
    if (source.which === "Custom") continue;
    const entry = cardEntry(source);
    const target = printCardFor(entry);
    if (!target) {
      unknown.push(entry);
      continue;
    }
    const group = groups.find((candidate) => candidate.card === target.card);
    if (group) group.entries.push(entry);
    else groups.push({ card: target.card, order: target.order, entries: [entry] });
  }
  groups.sort((a, b) => a.order - b.order);
  for (const group of groups) layoutGroup(group, writer);
  layoutCustomCards(custom, writer);
  for (const entry of unknown) layoutPlain(entry, writer);
  if (entries.length === 0) layoutNotice("No strap sheets chosen.", writer);
}

function layoutGroup({ card, entries }: Group, writer: PdfWriter): void {
  const widths = columnWidths(card);
  const xs = widths.reduce<number[]>((edges, width) => [...edges, (edges[edges.length - 1] ?? 0) + width], [0]);
  const tableWidth = xs[xs.length - 1] ?? SINGLE_WIDTH;
  const scale = Math.min(POINTS_PER_PIXEL, (PAGE_WIDTH - 2 * MARGIN_X) / tableWidth);
  const sheetBottom = (PAGE_HEIGHT - MARGIN_Y - BOTTOM_BORDER) / scale;
  const titleTop = sheetBottom - TITLE_ROW;
  const cardsBottom = titleTop - FOOTER_GAP;
  const bandHeight = card.rowHeights.reduce((sum, height) => sum + height, 0);
  const art = CARD_ART[card.which];

  const start = () => new PageCanvas({ left: (PAGE_WIDTH - tableWidth * scale) / 2, top: MARGIN_Y, scale, height: PAGE_HEIGHT });
  let canvas = start();
  let y = 0;
  const turn = () => {
    drawTitle(canvas, card.title, tableWidth, titleTop);
    writer.addPage(canvas);
    canvas = start();
    y = 0;
  };

  if (art) y = drawDrawing(canvas, writer, art, tableWidth) + DRAWING_GAP;
  for (let first = 0; first < entries.length; first += card.positions) {
    if (y + bandHeight > cardsBottom) turn();
    for (let position = 0; position < card.positions; position++) {
      const entry = entries[first + position];
      if (entry) drawCard(canvas, card, xs, y, position, entry);
    }
    y += bandHeight;
  }
  drawTitle(canvas, card.title, tableWidth, titleTop);
  writer.addPage(canvas);
}

function vectorCenterShift(art: CardArt): number {
  if (art.kind !== "vector") return 0;
  let minX = Infinity;
  let maxX = -Infinity;
  for (const item of art.items.slice(1)) {
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
  if (!Number.isFinite(minX) || !Number.isFinite(maxX)) return 0;
  return (minX + maxX) / 2 - art.w / 2;
}

function drawDrawing(canvas: PageCanvas, writer: PdfWriter, art: CardArt, tableWidth: number): number {
  const aspect = art.w / art.h;
  const width = Math.min(tableWidth * DRAWING_SHARE, DRAWING_MAX_HEIGHT * aspect);
  const height = width / aspect;
  const sx = width / art.w;
  const sy = height / art.h;
  if (art.kind === "vector") {
    const left = (tableWidth - width) / 2 - vectorCenterShift(art) * sx;
    canvas.paintItems(art.items, left, 0, sx, sy);
    return height;
  }
  const left = (tableWidth - width) / 2;
  canvas.image(writer.imageName(art), left, 0, width, height);
  return height;
}

function drawTitle(canvas: PageCanvas, title: string, tableWidth: number, top: number): void {
  let size = TITLE_SIZE;
  while (size > MIN_SIZE && textWidth(title, true, size) > tableWidth - 20) size -= 1;
  const width = textWidth(title, true, size);
  canvas.text(title, (tableWidth - width) / 2, top + TITLE_ROW - 2 - 0.21 * size, size, true, BLACK);
}

function drawCard(canvas: PageCanvas, card: PrintCard, xs: number[], top: number, position: number, entry: EntryPayload): void {
  const shift = position * card.posShift;
  const rowTop = (row: number) => top + card.rowHeights.slice(0, row).reduce((sum, height) => sum + height, 0);
  const cell = (row: number, column: number): Cell => {
    const left = xs[column + shift] ?? 0;
    const span = spanOf(card, row, column);
    return { x: left, y: rowTop(row), w: (xs[column + shift + span] ?? left) - left, h: card.rowHeights[row] ?? 0 };
  };
  const serialRow = card.serialCell[0];
  const serialColumn = card.serialCell[1];

  for (const labelRow of [0, 2]) {
    const labels = card.statics.filter(([row]) => row === labelRow);
    if (labels.length === 0) continue;
    const firstColumn = Math.min(...labels.map(([, column]) => column));
    const lastColumn = Math.max(...labels.map(([, column]) => column + spanOf(card, labelRow, column) - 1));
    const merged = new Set(labels.flatMap(([, column]) => Array.from({ length: spanOf(card, labelRow, column) - 1 }, (_, inner) => column + 1 + inner)));
    const metric = labels.filter(([, , text]) => METRIC_LABEL.test(text)).map(([, column]) => column);
    const metricStart = metric.length > 0 ? (card.metricFrom ?? Math.min(...metric)) : null;
    const inBlock = (column: number) => metricStart != null && column >= metricStart;

    for (const [, column] of labels) {
      const area = cell(labelRow, column);
      canvas.fillRect(area.x, area.y, area.w, area.h, inBlock(column) ? YELLOW : GRAY);
      if (inBlock(column)) {
        const below = cell(labelRow + 1, column);
        canvas.fillRect(below.x, below.y, below.w, below.h, YELLOW);
      }
    }
    const head = cell(labelRow, firstColumn);
    const tail = cell(labelRow, lastColumn);
    const left = head.x;
    const width = tail.x + tail.w - left;
    const height = head.h + (card.rowHeights[labelRow + 1] ?? 0);
    canvas.strokeRect(left, head.y, width, height, THIN);
    canvas.line(left, head.y + head.h, left + width, head.y + head.h, THIN);
    for (let column = firstColumn + 1; column <= lastColumn; column++) {
      if (merged.has(column)) continue;
      const x = cell(labelRow, column).x;
      canvas.line(x, head.y, x, head.y + height, THIN);
    }
    if (metricStart != null) {
      const block = cell(labelRow, metricStart);
      canvas.strokeRect(block.x, head.y, left + width - block.x, height, THICK);
    }
    for (const [, column, text] of labels) drawCell(canvas, cell(labelRow, column), text, LABEL_SIZE, "center");
  }

  for (const [row, column, from, also] of card.fields) {
    const text = [from, also].map((source) => (source == null ? "" : (entry.cells[source] ?? "").trim())).filter(Boolean).join(" ");
    if (text === "") continue;
    const isSerial = row === serialRow && column === serialColumn;
    const headerLabel = card.statics.find(([r, c]) => r === row - 1 && c === column)?.[2] ?? "";
    const valueSize = isSerial || LEGACY_VALUE_LABEL.test(headerLabel) ? VALUE_SIZE : LARGE_VALUE_SIZE;
    drawCell(canvas, cell(row, column), text, valueSize, isSerial ? "left" : "center");
  }
}

/**
 * Bold text in a cell. A value that is too wide wraps at spaces (two lines fit the value row) and
 * then shrinks; a serial stays on one line and shrinks. Whatever is left is clipped to the cell.
 */
function drawCell(canvas: PageCanvas, area: Cell, text: string, size: number, align: "center" | "left"): void {
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

/** An entry whose family the card tables do not know: its columns as a plain list, so nothing is lost. */
function layoutPlain(entry: EntryPayload, writer: PdfWriter): void {
  const canvas = new PageCanvas({ left: MARGIN_X, top: MARGIN_Y, scale: POINTS_PER_PIXEL, height: PAGE_HEIGHT });
  canvas.text(entry.which, 0, 28, 28, true, BLACK);
  Object.keys(entry.cells)
    .map(Number)
    .sort((a, b) => a - b)
    .forEach((column, index) => {
      const y = 60 + index * 24;
      canvas.text(`Column ${column}`, 0, y, 14, false, [0.35, 0.35, 0.35]);
      canvas.text(entry.cells[column] ?? "", 200, y, 16, true, BLACK);
    });
  writer.addPage(canvas);
}

function layoutNotice(text: string, writer: PdfWriter): void {
  const canvas = new PageCanvas({ left: MARGIN_X, top: MARGIN_Y, scale: POINTS_PER_PIXEL, height: PAGE_HEIGHT });
  canvas.text(text, 0, 28, 20, true, BLACK);
  writer.addPage(canvas);
}
