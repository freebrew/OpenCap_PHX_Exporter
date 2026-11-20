// Tiered OD/ID make-up torque lookup and deduplicated reference strip for StrapSheets PDF cards.
// Table data is generated from Z:\StrapSheets\design\torques.json into torques.gen.ts.

import type { PrintCard } from "./cardSpec";
import { parseMarks } from "./custom";
import type { EntryPayload } from "./entry";
import type { PageCanvas, Rgb } from "./pdfWriter";
import { TORQUE_DATA, type IdTier, type OdTier, type TorqueSpec } from "./torques.gen";

const BLACK: Rgb = [0, 0, 0];
const DARK_GRAY: Rgb = [0.32, 0.32, 0.32];
const MUTED: Rgb = [0.42, 0.42, 0.42];
const GRAY: Rgb = [0.953, 0.953, 0.953];
const SUB_HEADER: Rgb = [0.975, 0.975, 0.975];

const THIN = 1;
const HAIRLINE = 0.6;
const TITLE_BAR_H = 24;
const COL_HEAD_H = 20;
const ROW_H = 26;
const STRIP_GAP = 16;

export interface TorqueLookup {
  thread: string;
  canonical: string | null;
  odTier: OdTier | null;
  idTier: IdTier | null;
  assumedId: boolean;
  spec: TorqueSpec | null;
}

export interface TorqueRow {
  key: string;
  thread: string;
  odLabel: string | null;
  idLabel: string | null;
  assumedId: boolean;
  ftLbs: number | null;
  kNm: number | null;
}

const ALIAS_MAP: ReadonlyMap<string, string> = (() => {
  const map = new Map<string, string>();
  for (const [alias, canonical] of Object.entries(TORQUE_DATA.aliases)) {
    map.set(normalizeKey(alias), canonical);
  }
  return map;
})();

function normalizeKey(text: string): string {
  return text.trim().replace(/\s+/g, " ").toUpperCase();
}

/** Resolve any stored thread label (e.g. "4 1/2 IF [NC 50]", "DS 38", "4 1/2 XH") to its canonical chart key. */
export function canonicalThread(raw: string): string | null {
  const cleaned = normalizeKey(raw).replace(/^(?:BOX|PIN)\s+/, "");
  if (!cleaned) return null;
  return ALIAS_MAP.get(cleaned) ?? null;
}

/** Auto-detect inches vs mm from tool diameter magnitude: OD < 25 is inches, ID < 10 is inches. */
export function diameterToMm(raw: number, isId: boolean): number | null {
  if (!Number.isFinite(raw) || raw <= 0) return null;
  return (isId ? raw < 10 : raw < 25) ? raw * 25.4 : raw;
}

export function snapOdTier(odMm: number | null): OdTier | null {
  if (odMm == null || !Number.isFinite(odMm)) return null;
  return TORQUE_DATA.odTiers.find((tier) => odMm >= tier.minMm && odMm <= tier.maxMm) ?? null;
}

export function snapIdTier(idMm: number | null): IdTier | null {
  if (idMm == null || !Number.isFinite(idMm) || idMm > 105) return null;
  return TORQUE_DATA.idTiers.find((tier) => idMm <= tier.maxMm) ?? TORQUE_DATA.idTiers[TORQUE_DATA.idTiers.length - 1] ?? null;
}

/**
 * Look up a tiered make-up torque spec from a thread name and measured OD / ID (in mm or inches).
 * When ID is omitted (e.g. Motors, Orbits), falls back to the connection's nominal bore ID tier.
 * Proprietary or unknown threads return `spec: null` so the card prints "Verify mfr chart".
 */
export function lookupTorque(thread: string, odRaw: number | null, idRaw: number | null): TorqueLookup {
  const cleaned = thread.trim();
  const canonical = canonicalThread(cleaned);
  const odMm = odRaw != null ? diameterToMm(odRaw, false) : null;
  const idMm = idRaw != null ? diameterToMm(idRaw, true) : null;
  const odTier = snapOdTier(odMm);

  let idTier = snapIdTier(idMm);
  let assumedId = false;
  if (idTier == null && idMm == null && canonical) {
    const nominal = TORQUE_DATA.nominalId[canonical];
    if (nominal) {
      idTier = TORQUE_DATA.idTiers.find((tier) => tier.id === nominal) ?? null;
      assumedId = idTier != null;
    }
  }

  const spec = canonical && odTier && idTier
    ? TORQUE_DATA.specs[canonical]?.[odTier.id]?.[idTier.id] ?? null
    : null;

  return { thread: cleaned, canonical, odTier, idTier, assumedId, spec };
}

function parseNumber(raw: string | undefined): number | null {
  if (!raw) return null;
  const trimmed = raw.trim();
  if (!trimmed) return null;
  const value = Number(trimmed);
  return Number.isFinite(value) ? value : null;
}

function addTorqueRow(seen: Map<string, TorqueRow>, thread: string, odRaw: number | null, idRaw: number | null): void {
  const cleaned = thread.trim();
  if (!cleaned) return;
  const found = lookupTorque(cleaned, odRaw, idRaw);
  const odLabel = found.odTier?.label ?? null;
  const idLabel = found.idTier?.label ?? null;
  const key = `${normalizeKey(cleaned)}|${odLabel ?? ""}|${idLabel ?? ""}|${found.assumedId ? "1" : "0"}`;
  if (seen.has(key)) return;
  seen.set(key, {
    key,
    thread: cleaned,
    odLabel,
    idLabel,
    assumedId: found.assumedId,
    ftLbs: found.spec ? found.spec.ftLbs : null,
    kNm: found.spec ? found.spec.kNm : null,
  });
}

function fieldColumnForLabel(card: PrintCard, predicate: (label: string) => boolean): [number, number | undefined] | null {
  const stat = card.statics.find(([, , label]) => predicate(label));
  if (!stat) return null;
  const field = card.fields.find(([row, col]) => row === stat[0] + 1 && col === stat[1]);
  return field ? [field[2], field[3]] : null;
}

/**
 * Collect deduplicated make-up torque rows for a page of standard (or MWD) cards.
 * Uses the first and last lettered OD columns so middle stabilizer blades never skew connection OD.
 */
export function extractCardTorqueRows(card: PrintCard, entries: readonly EntryPayload[]): TorqueRow[] {
  const seen = new Map<string, TorqueRow>();
  const topField = fieldColumnForLabel(card, (label) => label === "TOP");
  const btmField = fieldColumnForLabel(card, (label) => label === "BTM");
  const topThreadCol = topField?.[1] ?? topField?.[0] ?? null;
  const btmThreadCol = btmField?.[1] ?? btmField?.[0] ?? null;
  if (topThreadCol == null && btmThreadCol == null) return [];

  const idField = fieldColumnForLabel(card, (label) => /^([A-Z] )?ID$/.test(label));
  const idCol = idField?.[0] ?? null;

  const letterOdCols = card.statics
    .filter(([, , label]) => /^[A-Z] OD$/.test(label))
    .map(([row, col]) => card.fields.find(([fr, fc]) => fr === row + 1 && fc === col)?.[2] ?? null)
    .filter((col): col is number => col != null);
  const maxOdField = fieldColumnForLabel(card, (label) => label === "MAX OD");
  const maxOdCol = maxOdField?.[0] ?? null;

  // On Subs (Floats, GAPs, X-Overs) B OD (first) is TOP/far-end and C OD (last) is BTM/pin-end;
  // on Motors, Orbits, MWD, Monels, etc., the first OD is lower/pin end and the last OD is upper end.
  const firstIsTop = card.which === "Subs" || card.which === "NM Flex Monel";

  for (const entry of entries) {
    const idRaw = idCol != null ? parseNumber(entry.cells[idCol]) : null;
    const firstOd = letterOdCols.length > 0 ? parseNumber(entry.cells[letterOdCols[0]!]) : null;
    const lastOd = letterOdCols.length > 1 ? parseNumber(entry.cells[letterOdCols[letterOdCols.length - 1]!]) : firstOd;
    const fallbackOd = maxOdCol != null ? parseNumber(entry.cells[maxOdCol]) : null;

    const topOd = (firstIsTop ? firstOd : lastOd) ?? (firstIsTop ? lastOd : firstOd) ?? fallbackOd;
    const btmOd = (firstIsTop ? lastOd : firstOd) ?? (firstIsTop ? firstOd : lastOd) ?? fallbackOd;

    if (topThreadCol != null) {
      const topThread = entry.cells[topThreadCol] ?? "";
      if (topThread.trim()) addTorqueRow(seen, topThread, topOd, idRaw);
    }
    if (btmThreadCol != null) {
      const btmThread = entry.cells[btmThreadCol] ?? "";
      if (btmThread.trim()) addTorqueRow(seen, btmThread, btmOd, idRaw);
    }
  }

  return [...seen.values()];
}

/** Collect deduplicated make-up torque rows for a page of Custom tubular cards. */
export function extractCustomTorqueRows(entries: readonly EntryPayload[]): TorqueRow[] {
  const seen = new Map<string, TorqueRow>();
  for (const entry of entries) {
    const idRaw = parseNumber(entry.cells.internalId);
    const marks = parseMarks(entry.cells.marksJson);
    const bodyOds = marks
      .filter((mark) => mark.kind !== "length" && mark.kind !== "pin" && mark.kind !== "spiral" && mark.kind !== "bitBox")
      .map((mark) => parseNumber(mark.value))
      .filter((value): value is number => value != null);
    const allOds = marks
      .filter((mark) => mark.kind !== "length" && mark.kind !== "pin")
      .map((mark) => parseNumber(mark.value))
      .filter((value): value is number => value != null);
    const odRaw = bodyOds.length > 0 ? Math.min(...bodyOds) : allOds.length > 0 ? Math.min(...allOds) : null;

    const topThread = entry.cells.threadUpper ?? "";
    const btmThread = entry.cells.threadLower ?? "";
    if (topThread.trim()) addTorqueRow(seen, topThread, odRaw, idRaw);
    if (btmThread.trim()) addTorqueRow(seen, btmThread, odRaw, idRaw);
  }
  return [...seen.values()];
}

/** Total vertical height (layout px) reserved above the family title, including gap. */
export function torqueStripReservedHeight(rowCount: number): number {
  if (rowCount <= 0) return 0;
  return TITLE_BAR_H + COL_HEAD_H + rowCount * ROW_H + STRIP_GAP;
}

function fitText(canvas: PageCanvas, text: string, bold: boolean, startSize: number, maxWidth: number): number {
  let size = startSize;
  while (size > 8 && canvas.measure(text, bold, size) > maxWidth) size -= 0.5;
  return size;
}

function tierText(canvas: PageCanvas, row: TorqueRow): string {
  if (row.odLabel && row.idLabel) {
    return row.assumedId
      ? `${row.odLabel} / ${row.idLabel} ${canvas.word("(nom.)")}`
      : `${row.odLabel} / ${row.idLabel}`;
  }
  if (row.odLabel) return row.odLabel;
  return canvas.word("Standard");
}

/**
 * Draw the deduplicated Make-Up Torque Reference table anchored at `bottomY` (just above the
 * family title). Returns the height drawn in layout pixels.
 */
export function drawTorqueStrip(canvas: PageCanvas, rows: readonly TorqueRow[], tableWidth: number, bottomY: number): number {
  if (rows.length === 0) return 0;

  const totalH = TITLE_BAR_H + COL_HEAD_H + rows.length * ROW_H;
  const topY = bottomY - totalH;
  const padX = 12;

  const colConnX = 0;
  const colTierX = Math.round(tableWidth * 0.28);
  const colFtX = Math.round(tableWidth * 0.62);
  const colKnmX = Math.round(tableWidth * 0.82);

  // Header band
  canvas.fillRect(0, topY, tableWidth, TITLE_BAR_H, GRAY);
  const heading = canvas.word("MAKE-UP TORQUE REFERENCE (TIERED OD / ID)");
  const disclaimer = canvas.word("API RP 7G REF ONLY - VERIFY MFR CHART");
  const headingSize = fitText(canvas, heading, true, 13, tableWidth * 0.58 - padX);
  const discSize = fitText(canvas, disclaimer, false, 11, tableWidth * 0.4 - padX);
  canvas.text(heading, padX, topY + 16.5, headingSize, true, BLACK);
  canvas.text(disclaimer, tableWidth - padX - canvas.measure(disclaimer, false, discSize), topY + 16, discSize, false, MUTED);

  // Sub-header row
  const subY = topY + TITLE_BAR_H;
  canvas.fillRect(0, subY, tableWidth, COL_HEAD_H, SUB_HEADER);
  canvas.line(0, subY, tableWidth, subY, THIN);

  const connHead = canvas.word("CONNECTION");
  const tierHead = canvas.word("NOMINAL SIZE TIER");
  const ftHead = canvas.word("FT-LBS");
  const knmHead = canvas.word("KN-M");
  canvas.text(connHead, colConnX + padX, subY + 14, fitText(canvas, connHead, true, 11, colTierX - 2 * padX), true, DARK_GRAY);
  canvas.text(tierHead, colTierX + padX, subY + 14, fitText(canvas, tierHead, true, 11, colFtX - colTierX - 2 * padX), true, DARK_GRAY);
  const ftHeadSize = fitText(canvas, ftHead, true, 11, colKnmX - colFtX - 2 * padX);
  canvas.text(ftHead, colKnmX - padX - canvas.measure(ftHead, true, ftHeadSize), subY + 14, ftHeadSize, true, DARK_GRAY);
  const knmHeadSize = fitText(canvas, knmHead, true, 11, tableWidth - colKnmX - 2 * padX);
  canvas.text(knmHead, tableWidth - padX - canvas.measure(knmHead, true, knmHeadSize), subY + 14, knmHeadSize, true, DARK_GRAY);

  // Data rows
  rows.forEach((row, index) => {
    const ry = subY + COL_HEAD_H + index * ROW_H;
    const rowBottom = ry + ROW_H;
    canvas.line(0, ry, tableWidth, ry, index === 0 ? THIN : HAIRLINE);

    const baseline = ry + 17.5;
    const connText = canvas.value(row.thread, false);
    const connSize = fitText(canvas, connText, true, 14, colTierX - 2 * padX);
    canvas.text(connText, colConnX + padX, baseline, connSize, true, BLACK);

    const tier = tierText(canvas, row);
    const tierSize = fitText(canvas, tier, false, 13, colFtX - colTierX - 2 * padX);
    canvas.text(tier, colTierX + padX, baseline, tierSize, false, DARK_GRAY);

    canvas.line(colTierX, ry, colTierX, rowBottom, HAIRLINE);
    canvas.line(colFtX, ry, colFtX, rowBottom, HAIRLINE);

    if (row.ftLbs != null && row.kNm != null) {
      canvas.line(colKnmX, ry, colKnmX, rowBottom, HAIRLINE);
      const ftText = `${row.ftLbs.toLocaleString("en-US")} ft-lbs`;
      const knmText = `${row.kNm} kN-m`;
      const ftSize = fitText(canvas, ftText, true, 14, colKnmX - colFtX - 2 * padX);
      const knmSize = fitText(canvas, knmText, true, 14, tableWidth - colKnmX - 2 * padX);
      canvas.text(ftText, colKnmX - padX - canvas.measure(ftText, true, ftSize), baseline, ftSize, true, BLACK);
      canvas.text(knmText, tableWidth - padX - canvas.measure(knmText, true, knmSize), baseline, knmSize, true, BLACK);
    } else {
      const noteText = canvas.word("Verify mfr chart");
      const spanW = tableWidth - colFtX;
      const noteSize = fitText(canvas, noteText, false, 13, spanW - 2 * padX);
      const noteW = canvas.measure(noteText, false, noteSize);
      canvas.text(noteText, colFtX + (spanW - noteW) / 2, baseline, noteSize, false, MUTED);
    }
  });

  // Sub-header vertical dividers + outer frame
  canvas.line(colTierX, subY, colTierX, subY + COL_HEAD_H, HAIRLINE);
  canvas.line(colFtX, subY, colFtX, subY + COL_HEAD_H, HAIRLINE);
  canvas.line(colKnmX, subY, colKnmX, subY + COL_HEAD_H, HAIRLINE);
  canvas.strokeRect(0, topY, tableWidth, totalH, THIN);

  return totalH;
}
