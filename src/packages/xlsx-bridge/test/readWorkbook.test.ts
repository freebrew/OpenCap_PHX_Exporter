import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";
import { readWorkbook } from "../src/readWorkbook.js";
import { ImportRejectedError } from "../src/types.js";
import { sharedFormulaFixture } from "./fixture.js";

describe("readWorkbook", () => {
  it("expands shared formulas, keeps array masters, and quarantines external refs", async () => {
    const mapped = await readWorkbook(sharedFormulaFixture());
    const sheet = mapped.workbook.sheets["sheet-1"];
    expect(sheet.name).toBe("Calc");
    expect(sheet.hidden).toBe(0);
    expect(sheet.freeze.ySplit).toBe(1);
    expect(sheet.cellData[0][0].v).toBe("Title");
    expect(sheet.cellData[0][2].f).toBe("=SUM(B1:B3)");
    expect(sheet.cellData[0][2].ref).toBe("C1");
    expect(sheet.cellData[1][2].f).toBe("=A2*B2");
    expect(sheet.cellData[2][2].f).toBe("=A3*B3");
    expect(sheet.cellData[2][3]?.f).toBeUndefined();
    expect(mapped.importLog.rejectedExternalRefs).toEqual([{ sheet: "Calc", cell: "D3" }]);
    expect(mapped.importLog.arrayFormulaCount).toBe(1);
    expect(mapped.importLog.sharedFormulaCount).toBe(2);
    expect(mapped.sidecar.protection[0]?.unlockedRanges).toContain("A2");
    expect(mapped.sidecar.protection[0]?.passwordPresent).toBe(true);
    expect(JSON.stringify(mapped.sidecar.protection)).not.toContain("secret");
    const answer = mapped.workbook.definedNames.find((name) => name.name === "Answer");
    const hidden = mapped.workbook.definedNames.find((name) => name.name === "HiddenConst");
    expect(answer?.formulaOrRefString).toBe("Calc!$C$1");
    expect(hidden?.formulaOrRefString).toBe("=1+2");
    expect(hidden?.hidden).toBe(true);
    expect(mapped.sidecar.names.find((name) => name.name === "Blob")?.supported).toBe(false);
    expect(mapped.sidecar.dataValidations[0]?.mapped).toBe(true);
    expect(mapped.sidecar.conditionalFormatting[0]?.mapped).toBe(true);
    const encoded = JSON.parse(JSON.stringify(mapped.sidecar)) as typeof mapped.sidecar;
    expect(encoded.names).toHaveLength(mapped.sidecar.names.length);
  });

  it("reads an ExcelJS workbook without committed binary fixtures", async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Inputs");
    sheet.getCell("A1").value = 2;
    sheet.getCell("B1").value = { formula: "A1*5", result: 10 };
    sheet.getCell("C1").value = "kept";
    const hidden = workbook.addWorksheet("Cache");
    hidden.state = "veryHidden";
    hidden.getCell("A1").value = 1;
    const bytes = await workbook.xlsx.writeBuffer();
    const mapped = await readWorkbook(new Uint8Array(bytes));
    const inputs = Object.values(mapped.workbook.sheets).find((item) => item.name === "Inputs");
    const cache = Object.values(mapped.workbook.sheets).find((item) => item.name === "Cache");
    expect(inputs?.cellData[0][0].v).toBe(2);
    expect(inputs?.cellData[0][1].f).toBe("=A1*5");
    expect(inputs?.cellData[0][2].v).toBe("kept");
    expect(cache?.hidden).toBe(2);
    expect(mapped.sidecar.veryHiddenSheets).toContain("Cache");
  });

  it("rejects a package that expands past the byte limit", async () => {
    await expect(readWorkbook(sharedFormulaFixture(), { limits: { maxUncompressedBytes: 32 } })).rejects.toBeInstanceOf(ImportRejectedError);
  });
});
