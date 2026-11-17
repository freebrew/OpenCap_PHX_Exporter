import { readFileSync, writeFileSync } from "node:fs";
import { readWorkbook, type CellFormulaRecord, type MappedWorkbook } from "../../xlsx-bridge/src/index.ts";
import { valuesMatch } from "./compare.js";
import { cellId, classifyFormulas } from "./classify.js";
import { calculateWorkbook, readCalculatedCell, reloadSnapshot } from "./headless.js";
import { oracleKey, readManifest, type OracleManifest } from "./oracle.js";

interface GateReport {
  sourceSha256: string;
  oracleSha256: string;
  sheets: number;
  formulaCells: number;
  udfDependent: number;
  evaluable: number;
  matches: number;
  engineGap: number;
  mapperBug: number;
  unclassified: number;
  rejectedExternal: number;
  mismatchRate: number;
  passed: boolean;
  udfCells: string[];
  engineGapExamples: string[];
  mapperBugExamples: string[];
  snapshotOk: boolean;
  snapshotDetail: string;
}

const EXAMPLE_CAP = 50;

function sheetIdByName(mapped: MappedWorkbook): Map<string, string> {
  return new Map(Object.values(mapped.workbook.sheets).map((sheet) => [sheet.name, sheet.id]));
}

function resource(workbook: MappedWorkbook["workbook"], name: string): string {
  return workbook.resources.find((item) => item.name === name)?.data ?? "";
}

function snapshotPreserves(mapped: MappedWorkbook, reloaded: MappedWorkbook["workbook"]): string {
  for (const sheetId of mapped.workbook.sheetOrder) {
    const before = mapped.workbook.sheets[sheetId];
    const after = reloaded.sheets[sheetId];
    if (!after) return `reload dropped ${before?.name ?? sheetId}`;
    if (before?.hidden !== after.hidden) return `${before?.name} hidden ${before?.hidden} -> ${after.hidden}`;
    if (JSON.stringify(before?.mergeData ?? []) !== JSON.stringify(after.mergeData ?? [])) return `${before?.name} merges changed`;
  }
  for (const formula of mapped.formulas) {
    const sheetId = sheetIdByName(mapped).get(formula.sheet);
    if (!sheetId) return `missing sheet ${formula.sheet}`;
    const before = mapped.workbook.sheets[sheetId]?.cellData[formula.row]?.[formula.col]?.f ?? "";
    const after = reloaded.sheets[sheetId]?.cellData[formula.row]?.[formula.col]?.f ?? "";
    if (normalizeFormula(before) !== normalizeFormula(after)) {
      return `${formula.sheet}!${formula.a1} formula ${before} -> ${after}`;
    }
  }
  const beforeNames = semanticNames(resource(mapped.workbook, "SHEET_DEFINED_NAME_PLUGIN"));
  const afterNames = semanticNames(resource(reloaded, "SHEET_DEFINED_NAME_PLUGIN"));
  if (beforeNames !== afterNames) return "defined names changed";
  const sidecar = JSON.stringify(mapped.sidecar);
  if (sidecar !== JSON.stringify(JSON.parse(sidecar))) return "sidecar json round-trip changed";
  return "";
}

function semanticNames(data: string): string {
  if (!data) return "";
  const parsed = JSON.parse(data) as Record<string, { name?: string; formulaOrRefString?: string; localSheetId?: string; hidden?: boolean }>;
  return JSON.stringify(Object.values(parsed).map((name) => ({
    name: name.name ?? "",
    formula: name.formulaOrRefString ?? "",
    sheet: name.localSheetId ?? "",
    hidden: Boolean(name.hidden),
  })).sort((a, b) => `${a.sheet}${a.name}`.localeCompare(`${b.sheet}${b.name}`)));
}

function normalizeFormula(formula: string): string {
  return formula.replace(/^=/, "").replace(/\s+/g, "");
}

export async function runParity(xlsmPath: string, oraclePath: string, timeoutMs = 120_000): Promise<GateReport> {
  const bytes = new Uint8Array(readFileSync(xlsmPath));
  process.stderr.write("mapping workbook\n");
  const mapped = await readWorkbook(bytes);
  process.stderr.write(`mapped ${mapped.importLog.sheetCount} sheets, ${mapped.formulas.length} formulas, depth ${mapped.importLog.observedDependencyDepth}\n`);
  const oracle = readManifest(oraclePath);
  if (oracle.sourceSha256.toLowerCase() !== mapped.sourceSha256.toLowerCase()) {
    throw new Error(`Oracle source hash ${oracle.sourceSha256} does not match workbook ${mapped.sourceSha256}.`);
  }
  process.stderr.write("classifying formulas\n");
  const classes = classifyFormulas(mapped.formulas, mapped.sidecar.names);
  process.stderr.write("calculating\n");
  const calculated = await calculateWorkbook(mapped.workbook, timeoutMs, mapped.importLog.unknownFunctions);
  process.stderr.write("reloading snapshot\n");
  const reloaded = reloadSnapshot(calculated);
  const ids = sheetIdByName(mapped);
  let udfDependent = 0;
  let evaluable = 0;
  let matches = 0;
  let engineGap = 0;
  let mapperBug = 0;
  let unclassified = 0;
  const udfCells: string[] = [];
  const engineGapExamples: string[] = [];
  const mapperBugExamples: string[] = [];

  for (const formula of mapped.formulas) {
    const kind = classes.get(cellId(formula));
    const label = `${formula.sheet}!${formula.a1}`;
    if (kind === "udf-dependent") {
      udfDependent += 1;
      udfCells.push(label);
      continue;
    }
    if (kind !== "evaluable") {
      unclassified += 1;
      continue;
    }
    evaluable += 1;
    const sheetId = ids.get(formula.sheet);
    const loaded = sheetId ? mapped.workbook.sheets[sheetId]?.cellData[formula.row]?.[formula.col]?.f : undefined;
    if (!loaded || normalizeFormula(loaded) !== normalizeFormula(formula.formula)) {
      mapperBug += 1;
      if (mapperBugExamples.length < EXAMPLE_CAP) mapperBugExamples.push(`${label} mapped formula missing`);
      continue;
    }
    const expected = oracle.cells[oracleKey(formula.sheet, formula.a1)];
    if (!expected) {
      mapperBug += 1;
      if (mapperBugExamples.length < EXAMPLE_CAP) mapperBugExamples.push(`${label} missing from sealed oracle`);
      continue;
    }
    const actual = readCalculatedCell(calculated, sheetId ?? "", formula.row, formula.col);
    if (valuesMatch(expected, actual)) {
      matches += 1;
      continue;
    }
    engineGap += 1;
    if (engineGapExamples.length < EXAMPLE_CAP) {
      engineGapExamples.push(`${label} expected ${JSON.stringify(expected)} actual ${JSON.stringify(actual)} formula ${formula.formula}`);
    }
  }

  const mismatchRate = evaluable === 0 ? 0 : (engineGap + mapperBug) / evaluable;
  const snapshotDetail = snapshotPreserves(mapped, reloaded);
  const snapshotOk = snapshotDetail === "";
  const passed = unclassified === 0 && snapshotOk && mismatchRate < 0.005;
  return {
    sourceSha256: mapped.sourceSha256,
    oracleSha256: oracle.sourceSha256,
    sheets: mapped.importLog.sheetCount,
    formulaCells: mapped.formulas.length,
    udfDependent,
    evaluable,
    matches,
    engineGap,
    mapperBug,
    unclassified,
    rejectedExternal: mapped.importLog.rejectedExternalRefs.length,
    mismatchRate,
    passed,
    udfCells,
    engineGapExamples,
    mapperBugExamples,
    snapshotOk,
    snapshotDetail,
  };
}

export function renderReport(report: GateReport, mappedLog?: string): string {
  const lines = [
    "# Demo formula parity",
    "",
    `Gate: ${report.passed ? "PASS" : "FAIL"}`,
    `Sheets: ${report.sheets}`,
    `Formula cells: ${report.formulaCells}`,
    `UDF-dependent (all): ${report.udfDependent}`,
    `Evaluable non-UDF: ${report.evaluable}`,
    `Matches: ${report.matches}`,
    `Engine gaps: ${report.engineGap}`,
    `Mapper bugs: ${report.mapperBug}`,
    `Unclassified: ${report.unclassified}`,
    `Rejected external refs: ${report.rejectedExternal}`,
    `Mismatch rate: ${(report.mismatchRate * 100).toFixed(4)}%`,
    `Snapshot round-trip: ${report.snapshotOk ? "ok" : `failed (${report.snapshotDetail})`}`,
    `Source SHA-256: ${report.sourceSha256}`,
    "",
    "## Engine gap examples",
    ...report.engineGapExamples.map((line) => `- ${line}`),
    "",
    "## Mapper bug examples",
    ...report.mapperBugExamples.map((line) => `- ${line}`),
    "",
    "## UDF-dependent cells",
    ...report.udfCells.map((line) => `- ${line}`),
    "",
  ];
  if (mappedLog) lines.push(mappedLog);
  return lines.join("\n");
}

function arg(name: string): string | undefined {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : undefined;
}

async function main(): Promise<void> {
  const xlsm = process.argv[2];
  const oracle = arg("--oracle");
  const out = arg("--out");
  if (!xlsm || xlsm.startsWith("--") || !oracle) {
    throw new Error("Usage: run.ts <path.xlsm> --oracle <manifest.json> [--out <report.md>]");
  }
  const timeout = Number(arg("--timeout") ?? "180000");
  const report = await runParity(xlsm, oracle, timeout);
  const text = renderReport(report);
  if (out) writeFileSync(out, text);
  process.stdout.write(`${text.split("## UDF-dependent")[0]}\n`);
  process.stdout.write(`UDF-dependent cells enumerated: ${report.udfCells.length}\n`);
  if (!report.passed) process.exitCode = 1;
}

const calledDirectly = process.argv[1] && import.meta.url.replace(/\\/g, "/").endsWith(process.argv[1].replace(/\\/g, "/"));
if (calledDirectly) {
  main().catch((error: unknown) => {
    process.stderr.write(`${error instanceof Error ? error.stack ?? error.message : String(error)}\n`);
    process.exitCode = 1;
  });
}

export type { CellFormulaRecord, OracleManifest };
