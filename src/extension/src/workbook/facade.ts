import type { CellWrite } from "../../../packages/workbook-commands/src/index";

export interface EditedRange {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
}

/** A floating shape's place on a sheet, like an Excel anchor. Rows and columns are 0-based; offsets are px. */
export interface SheetAnchor {
  col: number;
  colOff: number;
  row: number;
  rowOff: number;
  col2: number;
  colOff2: number;
  row2: number;
  rowOff2: number;
}

/** Where an anchored shape sits over the grid, in px from the grid's top-left. */
export interface SheetPlacement {
  left: number;
  top: number;
  width: number;
  height: number;
  zoom: number;
  /** The shape is hidden left of / above this line (the headers, or the frozen panes it scrolls under). */
  clipLeft: number;
  clipTop: number;
}

export interface WorkbookHost {
  load(snapshot: unknown): void;
  /** Take the workbook off the grid: the job now open has none, and the last job's must not stay under its name. */
  close(): void;
  snapshot(): unknown;
  onDirty(listener: () => void): () => void;
  setUdfNames(names: string[]): void;
  setCell(sheetName: string, a1: string, value: string | number | boolean | null): void;
  clearRange(sheetName: string, a1: string): void;
  getCell(sheetName: string, a1: string): unknown;
  getRawCell(sheetName: string, a1: string): unknown;
  /** Macro writes, in order. They do not raise onCellsEdited. */
  applyWrites(writes: CellWrite[]): void;
  flushCalculation(timeoutMs?: number): Promise<void>;
  activeSheet(): string;
  activateSheet(name: string): void;
  onActiveSheet(listener: (name: string) => void): () => void;
  /** Cells the user set by typing or from a dropdown. Rows and columns are 0-based. */
  onCellsEdited(listener: (sheetName: string, range: EditedRange) => void): () => void;
  /**
   * The selection moved. Rows and columns are 0-based. Sheet48's Worksheet_SelectionChange:
   * the refresh waits until the selection leaves F:G, so tabbing from Inc to Azm does not run it.
   */
  onSelectionChanged(listener: (sheetName: string, range: EditedRange) => void): () => void;
  /** Null unless `sheetName` is the sheet on screen. */
  placeOnSheet(sheetName: string, anchor: SheetAnchor): SheetPlacement | null;
  /** Scroll, zoom, resize, sheet switch or any command that can move cells. */
  onViewChange(listener: () => void): () => void;
  /** A command Excel cannot take back was refused, or it ran anyway. */
  onBlockedCommand(listener: (id: string, executed: boolean) => void): () => void;
  /** Univer's menu ribbon and formula bar. Hidden when the workbook host is created. */
  setMenusVisible(visible: boolean): void;
  dispose(): void;
}
