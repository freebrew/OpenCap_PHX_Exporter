/**
 * Formula parity with the real directional functions: recalculates every formula from scratch
 * and compares it with the result Excel saved in the workbook. Prints addresses, kinds and
 * formula text only, never cell values. Exit code 1 when anything differs.
 *
 * Usage: tsx src/packages/parity-harness/src/udfParity.ts "<workbook.xlsm>" [--examples N]
 */
import { readFileSync } from "node:fs";
import { readWorkbook } from "../../xlsx-bridge/src/index.ts";
import { legacyUdfEntries } from "../../../extension/src/workbook/formulaEntries.ts";
import { valuesMatch, type ActualCell } from "./compare.ts";
import { calculateWorkbook, readCalculatedCell } from "./headless.ts";

const path = process.argv[2];
if (!path) {
  process.stderr.write("Usage: udfParity.ts <workbook.xlsm> [--examples N]\n");
  process.exit(2);
}
const examplesArg = process.argv.indexOf("--examples");
const exampleLimit = examplesArg > 0 ? Number(process.argv[examplesArg + 1] ?? 5) : 5;

const started = Date.now();
const mapped = await readWorkbook(new Uint8Array(readFileSync(path)));
const udfs = new Set(mapped.importLog.unknownFunctions.map((name) => name.toUpperCase()));
const calculated = await calculateWorkbook(mapped.workbook, 300_000, [], legacyUdfEntries(mapped.importLog.unknownFunctions));
const sheetIds = new Map(Object.values(calculated.sheets).map((sheet) => [sheet.name, sheet.id]));

let matched = 0;
const bySheet = new Map<string, number>();
const byFunction = new Map<string, number>();
const examples: string[] = [];
for (const record of mapped.formulas) {
  const expected = expectedCell(record.cachedType, record.cachedValue);
  const actual = readCalculatedCell(calculated, sheetIds.get(record.sheet) ?? "", record.row, record.col);
  if (valuesMatch(expected, actual)) {
    matched += 1;
    continue;
  }
  const calls = [...record.formula.toUpperCase().matchAll(/([A-Z_][A-Z0-9_.]*)\(/g)].map((match) => match[1] ?? "");
  const head = calls.find((name) => udfs.has(name)) ?? calls[0] ?? "(reference)";
  bySheet.set(record.sheet, (bySheet.get(record.sheet) ?? 0) + 1);
  byFunction.set(head, (byFunction.get(head) ?? 0) + 1);
  if (examples.length < exampleLimit) examples.push(`${record.sheet}!${record.a1} [${expected.t} vs ${actual.t}] ${record.formula.slice(0, 120)}`);
}

const total = mapped.formulas.length;
const lines = [
  "============================================================",
  ` UDF PARITY | ${path.split(/[\\/]/).pop()} | ${Math.round((Date.now() - started) / 1000)} s`,
  "============================================================",
  ` Formulas   : ${total}`,
  ` Matched    : ${matched}`,
  ` Different  : ${total - matched}`,
];
if (total > matched) {
  lines.push(" By sheet   : " + [...bySheet].sort((a, b) => b[1] - a[1]).map(([name, count]) => `${name} ${count}`).join(", "));
  lines.push(" By function: " + [...byFunction].sort((a, b) => b[1] - a[1]).slice(0, 12).map(([name, count]) => `${name} ${count}`).join(", "));
  lines.push(" Examples   :", ...examples.map((line) => `   ${line}`));
}
lines.push(`RESULT: ${total === matched ? "PASS" : "DIFFERENT"}`);
process.stdout.write(`${lines.join("\n")}\n`);
process.exit(total === matched ? 0 : 1);

function expectedCell(type: string | undefined, value: unknown): ActualCell {
  if (value === undefined) return { t: "s", v: null };
  if (type === "e") return { t: "e", v: String(value) };
  if (type === "b") return { t: "b", v: Boolean(value) };
  if (type === "s" || type === "str") return { t: "s", v: String(value) };
  return { t: "n", v: value as number };
}
