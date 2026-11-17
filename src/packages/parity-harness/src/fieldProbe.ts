import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { readWorkbook } from "../../xlsx-bridge/src/index.ts";
import { calculateWorkbook, reloadSnapshot } from "./headless.ts";

const source = path.resolve(process.argv[2] ?? path.resolve(import.meta.dirname, "../../../../Slide Sheet - 35783.xlsm"));
if (path.basename(source) !== "Slide Sheet - 35783.xlsm") {
  throw new Error("The local probe only reads Slide Sheet - 35783.xlsm.");
}

const bytes = new Uint8Array(readFileSync(source));
const mapped = await readWorkbook(bytes);
const calculated = await calculateWorkbook(mapped.workbook, 180_000, mapped.importLog.unknownFunctions);
const reloaded = reloadSnapshot(calculated);
const directory = mkdtempSync(path.join(tmpdir(), "phx-field-probe-"));
const snapshotPath = path.join(directory, "preview.phx.json");
writeFileSync(snapshotPath, JSON.stringify({
  schemaVersion: 1,
  workbook: reloaded,
}));
const snapshotBytes = readFileSync(snapshotPath);
const snapshot = JSON.parse(snapshotBytes.toString("utf8")) as { workbook?: { sheetOrder?: unknown[] } };
rmSync(directory, { recursive: true, force: true });

const representative = firstNumericFormula(calculated);
console.log(`Field probe: read-only`);
console.log(`Sheets: ${mapped.importLog.sheetCount}`);
console.log(`Formulas: ${mapped.importLog.formulaCount}`);
console.log(`Very hidden: ${mapped.sidecar.veryHiddenSheets.length}`);
console.log(`Representative non-UDF recalc: ${representative}`);
console.log(`Snapshot reload sheets: ${snapshot.workbook?.sheetOrder?.length ?? 0}`);
console.log(`Disposable snapshot deleted: yes`);

interface ProbeCell {
  f?: string;
  v?: unknown;
}

interface ProbeSheet {
  name: string;
  cellData?: Record<string, Record<string, ProbeCell>>;
}

function firstNumericFormula(workbook: { sheets: Record<string, ProbeSheet> }): string {
  for (const sheet of Object.values(workbook.sheets)) {
    for (const row of Object.values(sheet.cellData ?? {})) {
      for (const cell of Object.values(row)) {
        if (!cell.f || mapped.importLog.unknownFunctions.some((name) => cell.f?.toUpperCase().includes(name))) continue;
        if (typeof cell.v === "number") return `${sheet.name} numeric`;
      }
    }
  }
  return "none";
}
