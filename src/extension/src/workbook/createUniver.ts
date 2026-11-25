import { commandBlocksHandoff, normalizeSnapshotFormulas } from "@phx/xlsx-bridge";
import { BorderStyleTypes, BorderType, CellValueType, HorizontalAlign, LocaleType, VerticalAlign, WrapStrategy, type IWorkbookData } from "@univerjs/core";
import { UniverDataValidationPlugin } from "@univerjs/data-validation";
import { UniverSheetsCorePreset } from "@univerjs/preset-sheets-core";
import sheetsCoreEnUS from "@univerjs/preset-sheets-core/locales/en-US";
import { createUniver } from "@univerjs/presets";
import { CalculationMode } from "@univerjs/sheets-formula";
import "@univerjs/sheets-formula/facade";
import { UniverSheetsConditionalFormattingPlugin } from "@univerjs/sheets-conditional-formatting";
import { UniverSheetsConditionalFormattingUIPlugin } from "@univerjs/sheets-conditional-formatting-ui";
import { UniverSheetsDataValidationPlugin } from "@univerjs/sheets-data-validation";
import { UniverSheetsDataValidationUIPlugin } from "@univerjs/sheets-data-validation-ui";
import { UniverSheetsHyperLinkPlugin } from "@univerjs/sheets-hyper-link";
import { UniverSheetsHyperLinkUIPlugin } from "@univerjs/sheets-hyper-link-ui";
import { UniverSheetsThreadCommentPlugin } from "@univerjs/sheets-thread-comment";
import { UniverSheetsThreadCommentUIPlugin } from "@univerjs/sheets-thread-comment-ui";
import { UniverThreadCommentPlugin } from "@univerjs/thread-comment";
import { UniverThreadCommentUIPlugin } from "@univerjs/thread-comment-ui";
import "@univerjs/sheets-numfmt/facade";
import "@univerjs/sheets-ui/facade";
import "@univerjs/preset-sheets-core/lib/index.css";
import type { CellWrite } from "../../../packages/workbook-commands/src/index";
import { legacyUdfEntries } from "./formulaEntries";
import type { EditedRange, SheetAnchor, SheetPlacement, WorkbookHost } from "./facade";
import { presentColumnFreeze, restoreColumnFreeze, type KeptColumnFreeze } from "./columnFreeze";
import { normalizeSlidesheetRowHeights } from "./slidesheetLayout";

interface SheetChrome {
  rowHeader?: { width?: number; hidden?: number };
  columnHeader?: { height?: number; hidden?: number };
}

function sheetChrome(snapshot: unknown): SheetChrome[] {
  const sheets = (snapshot as { sheets?: Record<string, SheetChrome> } | null)?.sheets;
  return sheets ? Object.values(sheets).filter((sheet): sheet is SheetChrome => sheet != null) : [];
}

/** Row numbers and A/B/C letters are preview chrome. The sheet's own headings stay. */
function hideSheetChrome(snapshot: unknown): void {
  for (const sheet of sheetChrome(snapshot)) {
    sheet.rowHeader = { width: sheet.rowHeader?.width ?? 46, hidden: 1 };
    sheet.columnHeader = { height: sheet.columnHeader?.height ?? 20, hidden: 1 };
  }
}

/** A saved snapshot keeps Excel's header gutter. The preview hides it again on load. */
function showSheetChrome(snapshot: unknown): void {
  for (const sheet of sheetChrome(snapshot)) {
    if (sheet.rowHeader) delete sheet.rowHeader.hidden;
    if (sheet.columnHeader) delete sheet.columnHeader.hidden;
  }
}

const SET_RANGE_VALUES = "sheet.command.set-range-values";
const CLEAR_CONTENT = "sheet.command.clear-selection-content";
const SET_STYLE = "sheet.command.set-style";
const SET_SELECTIONS = "sheet.operation.set-selections";

/** A macro string stays text: "1-5" is a downlink command, not 5 January. */
function typedCell(value: string | number | boolean): { v: string | number | boolean; t: CellValueType } {
  if (typeof value === "string") return { v: value, t: CellValueType.STRING };
  if (typeof value === "boolean") return { v: value, t: CellValueType.BOOLEAN };
  return { v: value, t: CellValueType.NUMBER };
}

/** The sheet canvas is the largest one Univer puts in the container. */
function sheetCanvas(container: HTMLElement): HTMLCanvasElement | null {
  let best: HTMLCanvasElement | null = null;
  for (const canvas of Array.from(container.querySelectorAll("canvas"))) {
    if (!best || canvas.clientWidth * canvas.clientHeight > best.clientWidth * best.clientHeight) best = canvas;
  }
  return best;
}

export function createWorkbookHost(container: HTMLElement): WorkbookHost {
  const worker = new Worker(new URL("./formula.worker.ts", import.meta.url), { type: "module" });
  worker.addEventListener("error", (event) => {
    console.error(`Formula worker: ${event.message}`);
  });
  const { univerAPI } = createUniver({
    locale: LocaleType.EN_US,
    locales: {
      [LocaleType.EN_US]: sheetsCoreEnUS,
    },
    presets: [
      UniverSheetsCorePreset({
        container,
        workerURL: worker,
        formula: {
          function: legacyUdfEntries() as never,
          // Saved values show at once; every formula is then recalculated with the ported functions,
          // because results saved by an older macro set (or with macros off) are not trusted.
          initialFormulaComputing: CalculationMode.FORCED,
        },
      }),
    ],
    plugins: [
      UniverDataValidationPlugin,
      UniverSheetsDataValidationPlugin,
      UniverSheetsDataValidationUIPlugin,
      UniverSheetsConditionalFormattingPlugin,
      UniverSheetsConditionalFormattingUIPlugin,
      UniverSheetsHyperLinkPlugin,
      UniverSheetsHyperLinkUIPlugin,
      UniverThreadCommentPlugin,
      UniverThreadCommentUIPlugin,
      UniverSheetsThreadCommentPlugin,
      UniverSheetsThreadCommentUIPlugin,
    ],
  });
  // The menu ribbon is Univer's toolbar part; the formula bar is drawn in its header part.
  function showMenus(visible: boolean): void {
    univerAPI
      .setUIVisible(univerAPI.Enum.BuiltInUIPart.TOOLBAR, visible)
      .setUIVisible(univerAPI.Enum.BuiltInUIPart.HEADER, visible);
  }
  showMenus(false);

  let unitId: string | null = null;
  let keptColumnFreezes: KeptColumnFreeze[] = [];
  let applying = 0;
  let guard = true;
  let blockedListener: ((id: string, executed: boolean) => void) | null = null;
  univerAPI.addEvent(univerAPI.Event.BeforeCommandExecute, (event) => {
    const command = event as { id?: string; cancel?: boolean };
    if (!guard || applying > 0 || !command.id || !commandBlocksHandoff(command.id)) return;
    command.cancel = true;
    blockedListener?.(command.id, false);
  });
  univerAPI.addEvent(univerAPI.Event.CommandExecuted, (event) => {
    const command = event as { id?: string };
    if (!guard || applying > 0 || !command.id || !commandBlocksHandoff(command.id)) return;
    blockedListener?.(command.id, true);
  });

  /** EnsureHiddenSheet: Cells.Clear an existing sheet, or add it hidden without leaving the sheet on screen. */
  function resetSheet(name: string): void {
    const workbook = univerAPI.getActiveWorkbook();
    if (!workbook) return;
    const existing = workbook.getSheetByName(name);
    if (existing) {
      existing.clear();
      return;
    }
    const active = workbook.getActiveSheet();
    const added = workbook.insertSheet(name);
    added.hideSheet();
    active.activate();
  }

  function applyWrite(write: CellWrite): void {
    if (write.kind === "sheet") {
      resetSheet(write.sheet);
      return;
    }
    const workbook = univerAPI.getActiveWorkbook();
    const sheet = workbook?.getSheetByName(write.sheet);
    if (!workbook || !sheet) return;
    if (write.kind === "width") {
      if (write.col > sheet.getMaxColumns()) sheet.setColumnCount(write.col);
      sheet.setColumnWidth(write.col - 1, write.px);
      return;
    }
    // Excel's grid is unbounded here; a Univer sheet only has the rows and columns it was saved with.
    const tableWidth = write.kind === "table" ? Math.max(1, ...write.values.map((line) => line.length)) : 1;
    const lastRow = "row2" in write ? write.row2 : write.kind === "table" ? write.row + Math.max(1, write.values.length) - 1 : write.row;
    const lastCol = "col2" in write ? write.col2 : write.col + tableWidth - 1;
    if (write.kind === "clear" || write.kind === "unmerge") {
      if (write.row > sheet.getMaxRows() || write.col > sheet.getMaxColumns()) return;
      write = { ...write, row2: Math.min(lastRow, sheet.getMaxRows()), col2: Math.min(lastCol, sheet.getMaxColumns()) };
    } else {
      if (lastRow > sheet.getMaxRows()) sheet.setRowCount(lastRow);
      if (lastCol > sheet.getMaxColumns()) sheet.setColumnCount(lastCol);
    }
    const area = (w: { row: number; col: number; row2: number; col2: number }) => sheet.getRange(w.row - 1, w.col - 1, w.row2 - w.row + 1, w.col2 - w.col + 1);
    const style = (w: { row: number; col: number; row2: number; col2: number }, type: string, value: unknown) => univerAPI.syncExecuteCommand(SET_STYLE, {
      unitId: workbook.getId(),
      subUnitId: sheet.getSheetId(),
      range: area(w).getRange(),
      style: { type, value },
    });
    switch (write.kind) {
      case "value":
        if (write.value == null || write.value === "") sheet.getRange(write.row - 1, write.col - 1).clearContent();
        else sheet.getRange(write.row - 1, write.col - 1).setValue(typedCell(write.value));
        return;
      case "table": {
        if (write.values.length === 0) return;
        const cells = write.values.map((line) => Array.from({ length: tableWidth }, (_, j) => {
          const value = line[j];
          return value == null || value === "" ? { v: null } : typedCell(value);
        }));
        sheet.getRange(write.row - 1, write.col - 1, cells.length, tableWidth).setValues(cells);
        return;
      }
      case "formula":
        sheet.getRange(write.row - 1, write.col - 1).setFormula(write.formula);
        return;
      case "clear":
        if (write.formats) area(write).clear();
        else area(write).clearContent();
        return;
      case "fill":
        style(write, "bg", write.fill ? { rgb: write.fill } : null);
        return;
      case "font":
        area(write).setFontColor(write.color);
        return;
      case "format":
        sheet.getRange(write.row - 1, write.col - 1).setNumberFormat(write.pattern);
        return;
      case "merge":
        area(write).merge();
        return;
      case "unmerge":
        if (area(write).isPartOfMerge()) area(write).breakApart();
        return;
      case "look": {
        const { bold, fontFamily, fontSize, hAlign, vAlign } = write.look;
        if (bold != null) style(write, "bl", bold ? 1 : 0);
        if (fontFamily) style(write, "ff", fontFamily);
        if (fontSize) style(write, "fs", fontSize);
        if (hAlign) style(write, "ht", hAlign === "left" ? HorizontalAlign.LEFT : hAlign === "center" ? HorizontalAlign.CENTER : HorizontalAlign.RIGHT);
        if (vAlign) style(write, "vt", VerticalAlign.MIDDLE);
        if (write.look.wrap === "clip") area(write).setWrapStrategy(WrapStrategy.CLIP);
        return;
      }
      case "border":
        area(write).setBorder(write.edge === "bottom" ? BorderType.BOTTOM : BorderType.OUTSIDE, write.line === "hair" ? BorderStyleTypes.HAIR : BorderStyleTypes.THIN, write.color);
        return;
      default: {
        const never: never = write;
        throw new Error(`Unhandled cell write ${String(never)}`);
      }
    }
  }

  function placeAnchor(sheetName: string, anchor: SheetAnchor): SheetPlacement | null {
    const sheet = univerAPI.getActiveWorkbook()?.getActiveSheet();
    const canvas = sheetCanvas(container);
    if (!sheet || !canvas || sheet.getSheetName() !== sheetName) return null;
    try {
      // getCellRect is in unzoomed sheet px and includes the row / column headers.
      const rect = (row: number, col: number) => sheet.getRange(row, col).getCellRect();
      const zoom = sheet.getZoom() || 1;
      const freeze = sheet.getFreeze();
      const frozenCols = Math.max(0, freeze.xSplit);
      const frozenRows = Math.max(0, freeze.ySplit);
      const origin = rect(0, 0);
      const scroll = sheet.getScrollState();
      const view = rect(scroll.sheetViewStartRow, scroll.sheetViewStartColumn);
      const scrolledCols = anchor.col >= frozenCols;
      const scrolledRows = anchor.row >= frozenRows;
      const scrollX = scrolledCols ? view.x - origin.x + scroll.offsetX : 0;
      const scrollY = scrolledRows ? view.y - origin.y + scroll.offsetY : 0;
      // An offset never reaches past its own cell, as in Excel.
      const corner = (row: number, col: number, rowOff: number, colOff: number) => {
        const cell = rect(row, col);
        const next = rect(row + 1, col + 1);
        return { x: cell.x + Math.min(colOff, next.x - cell.x), y: cell.y + Math.min(rowOff, next.y - cell.y) };
      };
      const a = corner(anchor.row, anchor.col, anchor.rowOff, anchor.colOff);
      const b = corner(anchor.row2, anchor.col2, anchor.rowOff2, anchor.colOff2);
      const canvasBox = canvas.getBoundingClientRect();
      const containerBox = container.getBoundingClientRect();
      const dx = canvasBox.left - containerBox.left;
      const dy = canvasBox.top - containerBox.top;
      const x1 = (a.x - scrollX) * zoom;
      const y1 = (a.y - scrollY) * zoom;
      const x2 = (b.x - scrollX) * zoom;
      const y2 = (b.y - scrollY) * zoom;
      const clipX = scrolledCols && frozenCols > 0 ? rect(0, frozenCols).x : origin.x;
      const clipY = scrolledRows && frozenRows > 0 ? rect(frozenRows, 0).y : origin.y;
      return {
        left: dx + x1,
        top: dy + y1,
        width: Math.max(0, x2 - x1),
        height: Math.max(0, y2 - y1),
        zoom,
        clipLeft: dx + clipX * zoom,
        clipTop: dy + clipY * zoom,
      };
    } catch {
      // The skeleton is not ready while a sheet is still loading.
      return null;
    }
  }

  return {
    load(snapshot) {
      guard = false;
      try {
        if (unitId) univerAPI.disposeUnit(unitId);
        const normalized = normalizeSnapshotFormulas(snapshot ?? {});
        normalizeSlidesheetRowHeights(normalized);
        const presented = presentColumnFreeze(normalized);
        keptColumnFreezes = presented.kept;
        hideSheetChrome(presented.view);
        const workbook = univerAPI.createWorkbook(presented.view as Partial<IWorkbookData>);
        unitId = workbook.getId();
        // The snapshot merge can still leave a frozen column. Clear it on the live sheet.
        for (const entry of keptColumnFreezes) {
          workbook.getSheetByName(entry.name)?.setFreeze({
            xSplit: 0,
            ySplit: entry.freeze.ySplit,
            startRow: entry.freeze.ySplit > 0 ? entry.freeze.startRow : -1,
            startColumn: -1,
          });
        }
      } finally {
        guard = true;
      }
    },
    close() {
      if (!unitId) return;
      guard = false;
      try {
        univerAPI.disposeUnit(unitId);
        unitId = null;
        keptColumnFreezes = [];
      } finally {
        guard = true;
      }
    },
    snapshot() {
      const saved = univerAPI.getActiveWorkbook()?.save() ?? null;
      if (!saved) return null;
      showSheetChrome(saved);
      return restoreColumnFreeze(saved, keptColumnFreezes);
    },
    onDirty(listener) {
      let timer = 0;
      const subscription = univerAPI.addEvent(univerAPI.Event.CommandExecuted, () => {
        window.clearTimeout(timer);
        timer = window.setTimeout(listener, 300);
      });
      return () => {
        window.clearTimeout(timer);
        subscription.dispose();
      };
    },
    setUdfNames() {
      // Stubs are constructed inside formula.worker.ts. Extra names stay #NAME?.
    },
    setCell(sheetName, a1, value) {
      const sheet = univerAPI.getActiveWorkbook()?.getSheetByName(sheetName);
      sheet?.getRange(a1).setValue(value ?? "");
    },
    clearRange(sheetName, a1) {
      const sheet = univerAPI.getActiveWorkbook()?.getSheetByName(sheetName);
      sheet?.getRange(a1).setValue("");
    },
    getCell(sheetName, a1) {
      const sheet = univerAPI.getActiveWorkbook()?.getSheetByName(sheetName);
      return sheet?.getRange(a1).getValue() ?? null;
    },
    getRawCell(sheetName, a1) {
      const sheet = univerAPI.getActiveWorkbook()?.getSheetByName(sheetName);
      return sheet?.getRange(a1).getRawValue() ?? null;
    },
    applyWrites(writes) {
      applying += 1;
      try {
        for (const write of writes) applyWrite(write);
      } finally {
        applying -= 1;
      }
    },
    async flushCalculation(timeoutMs = 60_000) {
      // Nothing to calculate, and Univer registers its formula commands only once a workbook exists:
      // asking before then throws "formula.mutation.set-trigger-formula-calculation-start is not registered".
      if (!univerAPI.getActiveWorkbook()) return;
      const formula = univerAPI.getFormula();
      formula.executeCalculation();
      await formula.onCalculationResultApplied(timeoutMs);
    },
    activeSheet() {
      return univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetName() ?? "";
    },
    activateSheet(name) {
      univerAPI.getActiveWorkbook()?.getSheetByName(name)?.activate();
    },
    onActiveSheet(listener) {
      const subscription = univerAPI.addEvent(univerAPI.Event.ActiveSheetChanged, (params) => {
        listener(params.activeSheet.getSheetName());
      });
      return () => subscription.dispose();
    },
    onSelectionChanged(listener) {
      // Every selection move (mouse, Tab / Enter, name box) runs this operation. It is watched
      // at the command bus, not through an FWorkbook: the host is created before the first
      // load(), and each load() disposes the unit and creates a new one.
      const subscription = univerAPI.addEvent(univerAPI.Event.CommandExecuted, (event) => {
        if (event.id !== SET_SELECTIONS) return;
        const params = (event.params ?? {}) as { unitId?: string; subUnitId?: string; selections?: Array<{ range?: EditedRange }> };
        const workbook = univerAPI.getActiveWorkbook();
        if (!workbook || (params.unitId && params.unitId !== workbook.getId())) return;
        const sheet = (params.subUnitId ? workbook.getSheetBySheetId(params.subUnitId) : null) ?? workbook.getActiveSheet();
        const range = params.selections?.[0]?.range;
        if (!sheet || !range) return;
        listener(sheet.getSheetName(), { startRow: range.startRow, endRow: range.endRow, startColumn: range.startColumn, endColumn: range.endColumn });
      });
      return () => subscription.dispose();
    },
    onCellsEdited(listener) {
      const subscription = univerAPI.addEvent(univerAPI.Event.CommandExecuted, (event) => {
        if (applying > 0 || (event.id !== SET_RANGE_VALUES && event.id !== CLEAR_CONTENT)) return;
        const params = (event.params ?? {}) as { subUnitId?: string; range?: EditedRange; ranges?: EditedRange[] };
        const workbook = univerAPI.getActiveWorkbook();
        const sheet = (params.subUnitId ? workbook?.getSheetBySheetId(params.subUnitId) : null) ?? workbook?.getActiveSheet();
        const ranges = params.ranges ?? [params.range ?? workbook?.getActiveRange()?.getRange()];
        if (!sheet) return;
        for (const range of ranges) if (range) listener(sheet.getSheetName(), range);
      });
      return () => subscription.dispose();
    },
    placeOnSheet(sheetName, anchor) {
      return placeAnchor(sheetName, anchor);
    },
    onViewChange(listener) {
      let frame = 0;
      const schedule = () => {
        if (frame) return;
        frame = window.requestAnimationFrame(() => {
          frame = 0;
          listener();
        });
      };
      const subscriptions = [
        univerAPI.addEvent(univerAPI.Event.Scroll, schedule),
        univerAPI.addEvent(univerAPI.Event.SheetZoomChanged, schedule),
        univerAPI.addEvent(univerAPI.Event.ActiveSheetChanged, schedule),
        univerAPI.addEvent(univerAPI.Event.CommandExecuted, schedule),
      ];
      const resize = new ResizeObserver(schedule);
      resize.observe(container);
      return () => {
        if (frame) window.cancelAnimationFrame(frame);
        resize.disconnect();
        for (const subscription of subscriptions) subscription.dispose();
      };
    },
    onBlockedCommand(listener) {
      blockedListener = listener;
      return () => {
        if (blockedListener === listener) blockedListener = null;
      };
    },
    setMenusVisible: showMenus,
    dispose() {
      univerAPI.dispose();
    },
  };
}
