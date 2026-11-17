import { readFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { readWorkbook } from "../src/readWorkbook.js";

const demoPath = path.resolve(import.meta.dirname, "../../../../Slide Sheet - Demo.xlsm");

describe("Demo workbook", () => {
  it("maps every sheet, shared formula, array master, name, and quarantines external refs", async () => {
    const mapped = await readWorkbook(new Uint8Array(readFileSync(demoPath)));
    expect(mapped.importLog.sheetCount).toBe(25);
    expect(mapped.importLog.formulaCount).toBeGreaterThan(20_000);
    expect(mapped.importLog.sharedFormulaCount).toBeGreaterThan(12_000);
    expect(mapped.importLog.arrayFormulaCount).toBe(16);
    expect(mapped.importLog.definedNameCount).toBe(987);
    expect(mapped.importLog.rejectedExternalRefs).toHaveLength(0);
    expect(mapped.sidecar.veryHiddenSheets).toContain("_OC_Job");
    const slidesheet = Object.values(mapped.workbook.sheets).find((sheet) => sheet.name === "Slidesheet");
    expect(slidesheet?.hidden).toBe(0);
    expect(mapped.sidecar.protection.some((sheet) => sheet.sheet === "Slidesheet" && sheet.protected)).toBe(true);
    expect(mapped.importLog.unknownFunctions.length).toBeGreaterThan(0);
    expect(JSON.stringify(mapped.sidecar.protection)).not.toMatch(/hashValue|saltValue|algorithmName/);
  }, 60_000);
});
