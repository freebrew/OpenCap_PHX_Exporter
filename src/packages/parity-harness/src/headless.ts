import { LocaleType, Univer } from "@univerjs/core";
import { FUniver } from "@univerjs/core/facade";
import { UniverDocsPlugin } from "@univerjs/docs";
import { UniverFormulaEnginePlugin } from "@univerjs/engine-formula";
import { UniverSheetsPlugin } from "@univerjs/sheets";
import { CalculationMode, UniverSheetsFormulaPlugin } from "@univerjs/sheets-formula";
import { UniverSheetsNumfmtPlugin } from "@univerjs/sheets-numfmt";
import "@univerjs/engine-formula/facade";
import "@univerjs/sheets/facade";
import "@univerjs/sheets-formula/facade";
import "@univerjs/sheets-numfmt/facade";
import type { UniverWorkbook } from "../../xlsx-bridge/src/index.ts";
import type { ActualCell } from "./compare.js";

export interface HeadlessSession {
  api: FUniver;
  dispose: () => void;
}

/** Custom functions for the engine, as `[class, name]` pairs (the extension's `legacyUdfEntries()`). */
export type FunctionEntries = unknown[];

export function createHeadless(mode: CalculationMode = CalculationMode.FORCED, functions?: FunctionEntries): HeadlessSession {
  const univer = new Univer({
    locale: LocaleType.EN_US,
    locales: {},
  });
  univer.registerPlugin(UniverFormulaEnginePlugin, functions ? { function: functions as never } : undefined);
  univer.registerPlugin(UniverDocsPlugin);
  univer.registerPlugin(UniverSheetsPlugin);
  univer.registerPlugin(UniverSheetsFormulaPlugin, {
    initialFormulaComputing: mode,
  });
  univer.registerPlugin(UniverSheetsNumfmtPlugin);
  const api = FUniver.newAPI(univer);
  return {
    api,
    dispose: () => univer.dispose(),
  };
}

function registerUdfStubs(api: FUniver, udfNames: string[]): void {
  const formula = api.getFormula();
  for (const name of udfNames) {
    formula.registerFunction(name, () => null, "Legacy VBA function. Phase 2 replaces this stub.");
  }
}

/** Excel's cached results ride along on formula cells; the gate must compare freshly computed values only. */
function withoutFormulaValues(workbook: UniverWorkbook): UniverWorkbook {
  const sheets: UniverWorkbook["sheets"] = {};
  for (const [id, sheet] of Object.entries(workbook.sheets)) {
    const cellData: typeof sheet.cellData = {};
    for (const [row, cells] of Object.entries(sheet.cellData)) {
      const out: (typeof sheet.cellData)[number] = {};
      for (const [col, cell] of Object.entries(cells)) {
        if (cell.f) {
          const { v: _cached, ...rest } = cell;
          out[Number(col)] = rest;
        } else {
          out[Number(col)] = cell;
        }
      }
      cellData[Number(row)] = out;
    }
    sheets[id] = { ...sheet, cellData };
  }
  return { ...workbook, sheets };
}

/** `functions` replaces the blank stubs with real implementations. */
export async function calculateWorkbook(workbook: UniverWorkbook, timeoutMs: number, udfNames: string[] = [], functions?: FunctionEntries): Promise<UniverWorkbook> {
  const session = createHeadless(CalculationMode.FORCED, functions);
  try {
    if (!functions) registerUdfStubs(session.api, udfNames);
    const created = session.api.createWorkbook(withoutFormulaValues(workbook));
    const formula = session.api.getFormula();
    formula.executeCalculation();
    await formula.onCalculationResultApplied(timeoutMs);
    return created.save() as UniverWorkbook;
  } finally {
    session.dispose();
  }
}

export function reloadSnapshot(workbook: UniverWorkbook): UniverWorkbook {
  const session = createHeadless(CalculationMode.NO_CALCULATION);
  try {
    return session.api.createWorkbook(workbook).save() as UniverWorkbook;
  } finally {
    session.dispose();
  }
}

export function readCalculatedCell(snapshot: UniverWorkbook, sheetId: string, row: number, col: number): ActualCell {
  const cell = snapshot.sheets[sheetId]?.cellData?.[row]?.[col];
  if (!cell) return { t: "s", v: null };
  const value = cell.v;
  if (value == null) return { t: "s", v: null };
  if (typeof value === "number") return { t: "n", v: value };
  if (typeof value === "boolean") return { t: "b", v: value };
  const text = String(value);
  if (/^#(NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA|SPILL!|CALC!)$/.test(text)) return { t: "e", v: text };
  return { t: "s", v: text };
}
