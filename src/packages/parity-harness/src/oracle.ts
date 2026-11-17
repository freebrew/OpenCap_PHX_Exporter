import { readFileSync, writeFileSync } from "node:fs";
import { readWorkbook, type CellFormulaRecord } from "../../xlsx-bridge/src/index.ts";

export interface OracleCell {
  t: "n" | "s" | "b" | "e";
  v: string | number | boolean | null;
}

export interface OracleManifest {
  schemaVersion: 1;
  sourceSha256: string;
  excelVersion: string;
  calculation: number;
  calculationState: number;
  sealedAt: string;
  formulaCells: number;
  cells: Record<string, OracleCell>;
}

export function oracleKey(sheet: string, a1: string): string {
  return `${sheet}!${a1}`;
}

function cacheToOracle(record: CellFormulaRecord): OracleCell {
  if (record.cachedType === "e") return { t: "e", v: String(record.cachedValue ?? "") };
  if (record.cachedType === "b" || typeof record.cachedValue === "boolean") return { t: "b", v: Boolean(record.cachedValue) };
  if (typeof record.cachedValue === "number") return { t: "n", v: record.cachedValue };
  if (record.cachedValue == null) return { t: "s", v: null };
  return { t: "s", v: String(record.cachedValue) };
}

export async function manifestFromCalculatedWorkbook(input: {
  calculatedPath: string;
  sourceSha256: string;
  excelVersion: string;
  calculation: number;
  calculationState: number;
  sealedAt: string;
}): Promise<OracleManifest> {
  const bytes = new Uint8Array(readFileSync(input.calculatedPath));
  const mapped = await readWorkbook(bytes);
  const cells: Record<string, OracleCell> = {};
  for (const formula of mapped.formulas) {
    cells[oracleKey(formula.sheet, formula.a1)] = cacheToOracle(formula);
  }
  return {
    schemaVersion: 1,
    sourceSha256: input.sourceSha256,
    excelVersion: input.excelVersion,
    calculation: input.calculation,
    calculationState: input.calculationState,
    sealedAt: input.sealedAt,
    formulaCells: mapped.formulas.length,
    cells,
  };
}

export function writeManifest(path: string, manifest: OracleManifest): void {
  writeFileSync(path, JSON.stringify(manifest));
}

export function readManifest(path: string): OracleManifest {
  return JSON.parse(readFileSync(path, "utf8")) as OracleManifest;
}
