import { describe, expect, it } from "vitest";
import type { UniverWorkbook } from "../../xlsx-bridge/src/index.ts";
import { calculateWorkbook, readCalculatedCell } from "../src/headless.ts";

function book(): UniverWorkbook {
  return {
    id: "probe",
    name: "Probe",
    appVersion: "0.1.0",
    locale: "enUS",
    styles: {},
    sheetOrder: ["s"],
    definedNames: [],
    resources: [],
    sheets: {
      s: {
        id: "s",
        name: "Data",
        hidden: 0,
        rowCount: 10,
        columnCount: 8,
        defaultColumnWidth: 80,
        defaultRowHeight: 20,
        freeze: { xSplit: 0, ySplit: 0, startRow: 0, startColumn: 0 },
        mergeData: [],
        rowData: {},
        columnData: {},
        showGridlines: 1,
        rowHeader: { width: 46 },
        columnHeader: { height: 20 },
        zoomRatio: 1,
        scrollTop: 0,
        scrollLeft: 0,
        rightToLeft: 0,
        cellData: {
          0: { 0: { v: 1 }, 1: { v: 10 } },
          1: { 0: { v: 2 }, 1: { v: 20 } },
          2: { 0: { v: 3 }, 1: { v: 30 } },
          3: {
            0: { f: "=SUM(A1:A3)" },
            1: { f: "=SUM(OFFSET(A1,0,0,3,1))" },
            2: { f: "=IFERROR(1/0,5)" },
            3: { f: "=INDEX(A1:A3,2)" },
            4: { f: "=MATCH(2,A1:A3,0)" },
            5: { f: "=VLOOKUP(2,A1:B3,2,FALSE)" },
            6: { f: "=ROUNDUP(1.1,0)" },
            7: { f: "=MOD(5,2)" },
          },
          4: {
            0: { f: "=ROWS(A1:A3)" },
            1: { f: '=LOOKUP(2,1/(A1:A3<>""),A1:A3)', ft: 2, ref: "B5" },
            2: { f: "=IF(ISNUMBER(A1),COS(((A1+A2)/2)*(PI()/180))*(B2-B1),\"\")" },
          },
        },
      },
    },
  };
}

describe("headless Univer", () => {
  it("calculates the Excel functions the slide sheet uses", async () => {
    const snapshot = await calculateWorkbook(book(), 20_000);
    const at = (row: number, col: number) => readCalculatedCell(snapshot, "s", row, col).v;
    expect(at(3, 0)).toBe(6);
    expect(at(3, 1)).toBe(6);
    expect(at(3, 2)).toBe(5);
    expect(at(3, 3)).toBe(2);
    expect(at(3, 4)).toBe(2);
    expect(at(3, 5)).toBe(20);
    expect(at(3, 6)).toBe(2);
    expect(at(3, 7)).toBe(1);
    expect(at(4, 0)).toBe(3);
    expect(at(4, 1)).toBe(3);
  }, 30_000);
});
