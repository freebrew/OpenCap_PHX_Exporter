export interface ImportLimits {
  maxUncompressedBytes: number;
  maxEntries: number;
  maxRows: number;
  maxColumns: number;
  maxFormulas: number;
  maxDependencyDepth: number;
  calculationTimeoutMs: number;
}

export const DEFAULT_LIMITS: ImportLimits = {
  maxUncompressedBytes: 200 * 1024 * 1024,
  maxEntries: 8_000,
  maxRows: 20_000,
  maxColumns: 512,
  maxFormulas: 100_000,
  maxDependencyDepth: 8_192,
  calculationTimeoutMs: 120_000,
};

export class ImportRejectedError extends Error {
  readonly code: string;

  constructor(code: string, message: string) {
    super(message);
    this.name = "ImportRejectedError";
    this.code = code;
  }
}

export type CellCache = string | number | boolean | null;

export interface CellFormulaRecord {
  sheet: string;
  a1: string;
  row: number;
  col: number;
  formula: string;
  arrayRef?: string;
  cachedValue?: CellCache;
  cachedType?: string;
}

export interface DefinedNameRecord {
  name: string;
  text: string;
  localSheetId?: number;
  hidden: boolean;
  builtin: boolean;
  supported: boolean;
  unsupportedReason?: string;
}

export interface ProtectionRecord {
  sheet: string;
  protected: boolean;
  options: Record<string, string>;
  passwordPresent: boolean;
  unlockedRanges: string[];
  hiddenFormulaRanges: string[];
}

export interface ValidationRecord {
  sheet: string;
  sqref: string;
  type: string;
  operator?: string;
  formula1?: string;
  formula2?: string;
  allowBlank: boolean;
  mapped: boolean;
  reason?: string;
}

export interface ConditionalFormatRecord {
  sheet: string;
  sqref: string;
  type: string;
  formula?: string;
  operator?: string;
  dxfId?: string;
  mapped: boolean;
  reason?: string;
}

export interface HyperlinkRecord {
  sheet: string;
  ref: string;
  target: string;
  confirmed: false;
  rejected: boolean;
  reason?: string;
}

export interface SheetLog {
  name: string;
  state: "visible" | "hidden" | "veryHidden";
  formulas: number;
  rows: number;
  cols: number;
}

export interface ImportLog {
  sheetCount: number;
  formulaCount: number;
  sharedFormulaCount: number;
  arrayFormulaCount: number;
  definedNameCount: number;
  supportedNameCount: number;
  unknownFunctions: string[];
  rejectedExternalRefs: Array<{ sheet: string; cell: string }>;
  dropped: {
    vba: boolean;
    drawings: number;
    charts: number;
    comments: number;
    formControls: number;
  };
  unsupportedConditionalFormats: number;
  unsupportedValidations: number;
  excelJsFormulaDisagreements: number;
  observedDependencyDepth: number;
  warnings: string[];
  perSheet: SheetLog[];
  macroButtonCount: number;
}

export interface UniverCell {
  v?: string | number | boolean;
  f?: string;
  /** 0 normal, 1 shared, 2 array, 3 data table. */
  ft?: number;
  t?: number;
  s?: string;
  ref?: string;
  custom?: Record<string, unknown>;
}

export interface UniverRange {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
}

export interface UniverSheet {
  id: string;
  name: string;
  tabColor?: string;
  hidden: 0 | 1 | 2;
  rowCount: number;
  columnCount: number;
  defaultColumnWidth: number;
  defaultRowHeight: number;
  freeze: { xSplit: number; ySplit: number; startRow: number; startColumn: number };
  mergeData: UniverRange[];
  cellData: Record<number, Record<number, UniverCell>>;
  rowData: Record<number, { h?: number; hd?: 0 | 1 }>;
  columnData: Record<number, { w?: number; hd?: 0 | 1 }>;
  showGridlines: 0 | 1;
  rowHeader: { width: number };
  columnHeader: { height: number };
  zoomRatio: number;
  scrollTop: number;
  scrollLeft: number;
  rightToLeft: 0 | 1;
}

export interface UniverDefinedName {
  id: string;
  name: string;
  formulaOrRefString: string;
  localSheetId?: string;
  hidden?: boolean;
}

export interface UniverWorkbook {
  id: string;
  name: string;
  appVersion: string;
  locale: string;
  styles: Record<string, Record<string, unknown>>;
  sheetOrder: string[];
  sheets: Record<string, UniverSheet>;
  definedNames: UniverDefinedName[];
  resources: Array<{ name: string; data: string }>;
}

export interface MacroButtonRecord {
  sheet: string;
  id: string;
  caption: string;
  macro: string;
  anchor: string;
}

export interface WorkbookSidecar {
  veryHiddenSheets: string[];
  names: DefinedNameRecord[];
  protection: ProtectionRecord[];
  dataValidations: ValidationRecord[];
  conditionalFormatting: ConditionalFormatRecord[];
  hyperlinks: HyperlinkRecord[];
  macroButtons?: MacroButtonRecord[];
  styleVersion?: number;
}

export interface MappedWorkbook {
  workbook: UniverWorkbook;
  sidecar: WorkbookSidecar;
  formulas: CellFormulaRecord[];
  importLog: ImportLog;
  sourceSha256: string;
}

export interface ReadOptions {
  limits?: Partial<ImportLimits>;
  /** When true, a workbook with external references is rejected instead of quarantining those cells. */
  rejectExternalReferences?: boolean;
}
