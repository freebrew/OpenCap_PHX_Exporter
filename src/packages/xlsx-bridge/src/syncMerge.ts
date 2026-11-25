// Visible-tab user data only. Formulas, hidden sheets, styles and layout stay put.
// Numbers compare by value (relative 1e-9) so a JS 17-digit serial and an Excel
// double of the same time or depth are the same cell. "00" and "2-4" stay text.

import { parseA1 } from "./a1.js";
import type { CanonBook, CanonCell, CanonSheet, CellKind, HandoffOp } from "./handoff.js";

const RELATIVE = 1e-9;
const PLAIN_NUMBER = /^-?(?:0|[1-9]\d*)(?:\.\d+)?$/;

export type ScalarKind = Exclude<CellKind, "formula">;

export interface ScalarCell {
  kind: ScalarKind;
  value: string;
}

export interface SyncConflict {
  sheet: string;
  a1: string;
  base: ScalarCell;
  excel: ScalarCell;
  browser: ScalarCell;
}

export type CellOp = Extract<HandoffOp, { kind: "cell" }>;

export interface VisibleMerge {
  /** Changed in Excel only. The browser takes these. */
  fromExcel: CellOp[];
  /** Changed in the browser only. They stay, and the next export sends them. */
  keepBrowser: CellOp[];
  /** Both sides made the same change. Nothing to write. */
  converged: CellOp[];
  /** Both sides changed, and not to the same value. */
  conflicts: SyncConflict[];
}

/** One-way Browser -> Excel tabs. Excel never overwrites or conflicts with browser values here. */
const BROWSER_AUTHORITATIVE_SHEETS = new Set(["setup", "slidesheet"]);

/** Code-painted / computed Data tables rebuilt from _OC_* / _TH_Hours / _OC_Survey / _OC_AC / Slidesheet. Never diffed as user edits. */
const DERIVED_RANGES: ReadonlyArray<{ sheet: string; r1: number; r2: number; c1: number; c2: number }> = [
  { sheet: "Data", r1: 13, r2: 15, c1: 5, c2: 6 }, // E13:F15 Wellbore Position (MDL_WellborePos / planGauge)
  { sheet: "Data", r1: 35, r2: 42, c1: 2, c2: 6 }, // B35:F42 AC Info & Concerns (MDL_ACImport / acImport)
  { sheet: "Data", r1: 45, r2: 55, c1: 2, c2: 2 }, // B45:B55 Motors on Location from OpenCap (MDL_Setup / openCap)
  { sheet: "Data", r1: 30, r2: 41, c1: 15, c2: 16 }, // O30:P41 active-BHA 3rd-party serials (MDL_ToolHours)
  { sheet: "Data", r1: 31, r2: 41, c1: 18, c2: 19 }, // R31:S41 3rd-party current/total hours (Q stays user-typed)
  { sheet: "Data", r1: 44, r2: 55, c1: 15, c2: 16 }, // O44:P55 active-BHA motor serials (MDL_ToolHours)
  { sheet: "Data", r1: 45, r2: 55, c1: 18, c2: 19 }, // R45:S55 motor current/total hours (Q stays user-typed)
  { sheet: "Data", r1: 54, r2: 54, c1: 9, c2: 13 }, // I54:M54 Planned TD from _OC_Survey (MDL_TDCalc / planImport)
  { sheet: "Data", r1: 2, r2: 400, c1: 24, c2: 28 }, // X2:AB400 BHA serial hours (Module23 / bhaSelect)
];

/** Active-BHA-scoped cells on Data. When one side has a higher BHA # on Data!H3, the newer BHA owns this entire block atomically. */
const ACTIVE_BHA_RANGES: ReadonlyArray<{ sheet: string; r1: number; r2: number; c1: number; c2: number }> = [
  { sheet: "Data", r1: 3, r2: 3, c1: 8, c2: 8 }, // H3 Active BHA #
  { sheet: "Data", r1: 22, r2: 22, c1: 3, c2: 3 }, // C22 Active Motor S/N
  { sheet: "Data", r1: 4, r2: 22, c1: 8, c2: 17 }, // H4:Q22 Day Roll table (formula cols K & O are skipped before this check)
];

/** OpenCap / FieldCap cells on bidirectional tabs when both sides share the same BHA #. A conflict follows the newer pull. */
const MACHINE_RANGES: ReadonlyArray<{ sheet: string; r1: number; r2: number; c1: number; c2: number }> = [
  { sheet: "Data", r1: 4, r2: 22, c1: 12, c2: 14 }, // L4:N22 slide metres & slide time
  { sheet: "Data", r1: 22, r2: 22, c1: 3, c2: 3 }, // C22 active motor S/N
];

export function isBrowserAuthoritativeSheet(sheet: string): boolean {
  return BROWSER_AUTHORITATIVE_SHEETS.has(sheet.trim().toLowerCase());
}

function inRanges(ranges: ReadonlyArray<{ sheet: string; r1: number; r2: number; c1: number; c2: number }>, sheet: string, a1: string): boolean {
  let row = 0;
  let col = 0;
  try {
    const pos = parseA1(a1);
    row = pos.row + 1;
    col = pos.col + 1;
  } catch {
    return false;
  }
  return ranges.some((range) => range.sheet === sheet && row >= range.r1 && row <= range.r2 && col >= range.c1 && col <= range.c2);
}

export function isDerivedCell(sheet: string, a1: string): boolean {
  return inRanges(DERIVED_RANGES, sheet, a1);
}

export function isActiveBhaCell(sheet: string, a1: string): boolean {
  return inRanges(ACTIVE_BHA_RANGES, sheet, a1);
}

function activeBhaNumber(sheet: CanonSheet | undefined): number {
  const cell = sheet?.cells["H3"];
  if (!cell || cell.k === "formula") return 0;
  const n = numericValue(cell.k, cell.v);
  return n != null && n > 0 ? Math.round(n) : 0;
}

export function numbersClose(left: number, right: number): boolean {
  const a = Object.is(left, -0) ? 0 : left;
  const b = Object.is(right, -0) ? 0 : right;
  const scale = Math.max(Math.abs(a), Math.abs(b), 1);
  return Math.abs(a - b) <= RELATIVE * scale;
}

function numericValue(kind: CellKind, value: string): number | null {
  if (kind === "number") {
    const n = Number(value);
    return Number.isFinite(n) ? n : null;
  }
  if (kind === "text" && PLAIN_NUMBER.test(value)) {
    const n = Number(value);
    return Number.isFinite(n) ? n : null;
  }
  return null;
}

function scalarOf(cell: CanonCell): ScalarCell {
  if (cell.k === "formula" || (cell.k === "text" && cell.v === "")) return { kind: "blank", value: "" };
  return { kind: cell.k, value: cell.v };
}

/** True when two non-formula cells are the same value, including 1.75 and 1.750. */
export function scalarEqual(left: CanonCell, right: CanonCell): boolean {
  if (left.k === "formula" || right.k === "formula") return false;
  const ln = numericValue(left.k, left.v);
  const rn = numericValue(right.k, right.v);
  if (ln != null && rn != null) return numbersClose(ln, rn);
  const a = scalarOf(left);
  const b = scalarOf(right);
  if (a.kind !== b.kind) return false;
  if (a.kind === "blank") return true;
  if (a.kind === "bool") return a.value.toUpperCase() === b.value.toUpperCase();
  return a.value === b.value;
}

export function machineWrittenCell(sheet: string, a1: string): boolean {
  return inRanges(MACHINE_RANGES, sheet, a1);
}

/** Excel stays the authority for a typed cell. Machine cells follow the newer FieldCap pull. */
export function defaultConflictChoice(conflict: SyncConflict, preferBrowserMachine: boolean): "excel" | "browser" {
  if (machineWrittenCell(conflict.sheet, conflict.a1)) return preferBrowserMachine ? "browser" : "excel";
  return "excel";
}

function blankCell(): CanonCell {
  return { k: "blank", v: "" };
}

function cellAt(sheet: CanonSheet | undefined, a1: string): CanonCell {
  return sheet?.cells[a1] ?? blankCell();
}

function cellOp(sheet: string, a1: string, base: CanonCell, next: CanonCell): CellOp {
  const from = scalarOf(base);
  const to = scalarOf(next);
  return {
    kind: "cell",
    sheet,
    a1,
    baseKind: from.kind,
    baseValue: from.value,
    nextKind: to.kind,
    nextValue: to.value,
  };
}

function compareA1(a: string, b: string): number {
  const left = a.match(/^([A-Z]+)(\d+)$/);
  const right = b.match(/^([A-Z]+)(\d+)$/);
  if (!left || !right) return a.localeCompare(b);
  const row = Number(left[2]) - Number(right[2]);
  if (row !== 0) return row;
  return left[1].localeCompare(right[1]);
}

function visiblePair(base: CanonBook, live: CanonBook): Array<{ name: string; prior: CanonSheet; live: CanonSheet }> {
  const priorByName = new Map(base.sheets.map((sheet) => [sheet.name, sheet]));
  const pairs: Array<{ name: string; prior: CanonSheet; live: CanonSheet }> = [];
  for (const sheet of live.sheets) {
    const prior = priorByName.get(sheet.name);
    if (!prior || prior.hidden !== 0 || sheet.hidden !== 0) continue;
    pairs.push({ name: sheet.name, prior, live: sheet });
  }
  return pairs;
}

/** Cell edits on sheets that are visible on both sides. Formulas and derived OpenCap tables are left alone. */
export function diffVisibleUserData(base: CanonBook, live: CanonBook): CellOp[] {
  const ops: CellOp[] = [];
  for (const pair of visiblePair(base, live)) {
    const addresses = new Set([...Object.keys(pair.prior.cells), ...Object.keys(pair.live.cells)]);
    for (const a1 of [...addresses].sort(compareA1)) {
      if (isDerivedCell(pair.name, a1)) continue;
      const baseCell = cellAt(pair.prior, a1);
      const liveCell = cellAt(pair.live, a1);
      if (baseCell.k === "formula" || liveCell.k === "formula") continue;
      if (scalarEqual(baseCell, liveCell)) continue;
      ops.push(cellOp(pair.name, a1, baseCell, liveCell));
    }
  }
  return ops;
}

/**
 * Three books, one classification per changed scalar on a sheet that is visible in all three.
 * Setup and Slidesheet are strictly Browser-authoritative (one-way Browser -> Excel): Excel
 * changes there never pull into the browser or raise conflicts.
 * Data derived tables (Wellbore Position, AC Concerns, Motors on Location, Tool/Motor Hours,
 * Planned TD, BHA Serial Hours) are skipped and rebuilt by code hooks.
 * When one side has a higher active BHA # on Data!H3, that newer BHA owns the entire
 * BHA-scoped Data block (H3, C22, H4:Q22) atomically with zero conflicts or mixed rows.
 */
export function mergeVisibleUserData(base: CanonBook, browser: CanonBook, excel: CanonBook): VisibleMerge {
  const result: VisibleMerge = { fromExcel: [], keepBrowser: [], converged: [], conflicts: [] };
  const browserByName = new Map(browser.sheets.map((sheet) => [sheet.name, sheet]));
  const excelByName = new Map(excel.sheets.map((sheet) => [sheet.name, sheet]));
  const browserBha = activeBhaNumber(browserByName.get("Data"));
  const excelBha = activeBhaNumber(excelByName.get("Data"));
  const newerBhaSide: "excel" | "browser" | null =
    excelBha > 0 && browserBha > 0 && excelBha !== browserBha
      ? (excelBha > browserBha ? "excel" : "browser")
      : null;

  for (const prior of base.sheets) {
    const browserSheet = browserByName.get(prior.name);
    const excelSheet = excelByName.get(prior.name);
    if (!browserSheet || !excelSheet) continue;
    if (prior.hidden !== 0 || browserSheet.hidden !== 0 || excelSheet.hidden !== 0) continue;
    const browserWins = isBrowserAuthoritativeSheet(prior.name);
    const addresses = new Set([
      ...Object.keys(prior.cells),
      ...Object.keys(browserSheet.cells),
      ...Object.keys(excelSheet.cells),
    ]);
    for (const a1 of [...addresses].sort(compareA1)) {
      if (isDerivedCell(prior.name, a1)) continue;
      const baseCell = cellAt(prior, a1);
      const browserCell = cellAt(browserSheet, a1);
      const excelCell = cellAt(excelSheet, a1);
      if (baseCell.k === "formula" || browserCell.k === "formula" || excelCell.k === "formula") continue;
      if (browserWins) {
        if (!scalarEqual(browserCell, excelCell)) {
          result.keepBrowser.push(cellOp(prior.name, a1, excelCell, browserCell));
        } else if (!scalarEqual(baseCell, browserCell)) {
          result.converged.push(cellOp(prior.name, a1, baseCell, browserCell));
        }
        continue;
      }
      if (newerBhaSide && isActiveBhaCell(prior.name, a1)) {
        if (newerBhaSide === "excel") {
          if (!scalarEqual(excelCell, browserCell)) {
            result.fromExcel.push(cellOp(prior.name, a1, browserCell, excelCell));
          } else if (!scalarEqual(baseCell, excelCell)) {
            result.converged.push(cellOp(prior.name, a1, baseCell, excelCell));
          }
        } else {
          if (!scalarEqual(browserCell, excelCell)) {
            result.keepBrowser.push(cellOp(prior.name, a1, excelCell, browserCell));
          } else if (!scalarEqual(baseCell, browserCell)) {
            result.converged.push(cellOp(prior.name, a1, baseCell, browserCell));
          }
        }
        continue;
      }
      const browserChanged = !scalarEqual(baseCell, browserCell);
      const excelChanged = !scalarEqual(baseCell, excelCell);
      if (!browserChanged && !excelChanged) continue;
      if (browserChanged && excelChanged && scalarEqual(browserCell, excelCell)) {
        result.converged.push(cellOp(prior.name, a1, baseCell, browserCell));
        continue;
      }
      if (browserChanged && excelChanged) {
        result.conflicts.push({
          sheet: prior.name,
          a1,
          base: scalarOf(baseCell),
          excel: scalarOf(excelCell),
          browser: scalarOf(browserCell),
        });
        continue;
      }
      if (excelChanged) result.fromExcel.push(cellOp(prior.name, a1, baseCell, excelCell));
      else result.keepBrowser.push(cellOp(prior.name, a1, baseCell, browserCell));
    }
  }
  return result;
}
