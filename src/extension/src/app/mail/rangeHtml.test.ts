import { describe, expect, it } from "vitest";
import type { SnapshotBook } from "../../../../packages/workbook-commands/src/index";
import { formatExcelNumber, rangeToHtml } from "./rangeHtml";

describe("email number formats", () => {
  it("prints the same text Excel shows on Data", () => {
    expect(formatExcelNumber(3186, '######.00"m"')).toBe("3186.00m");
    expect(formatExcelNumber(671, '0"m"')).toBe("671m");
    expect(formatExcelNumber(0, '#"Hrs"')).toBe("0Hrs");
    expect(formatExcelNumber(28672.5, '"$"#,##0.00')).toBe("$28,672.50");
    expect(formatExcelNumber(251663.75, '"$"#,##0.00')).toBe("$251,663.75");
    expect(formatExcelNumber(4706.9, "####.00\\°")).toBe("4706.90°");
    expect(formatExcelNumber(0.523254, '####.00"°/30m"')).toBe("0.52°/30m");
    expect(formatExcelNumber(4.9, '####.0"°/30m"')).toBe("4.9°/30m");
    expect(formatExcelNumber(98, '##.####"°C"')).toBe("98°C");
    expect(formatExcelNumber(2.2, '###.00"m³/min"')).toBe("2.20m³/min");
    expect(formatExcelNumber(50, '####.0"rpm"')).toBe("50.0rpm");
    expect(formatExcelNumber(349.8, '####.0"rpm"')).toBe("349.8rpm");
    expect(formatExcelNumber(25.75, '#####.00"Hrs";\\-#####.00"Hrs";')).toBe("25.75Hrs");
    expect(formatExcelNumber(0, '#####.00"Hrs";\\-#####.00"Hrs";')).toBe("");
    expect(formatExcelNumber(-2, '#####.00"Hrs";\\-#####.00"Hrs";')).toBe("-2.00Hrs");
    expect(formatExcelNumber(0.185, "0.00%")).toBe("18.50%");
    expect(formatExcelNumber(1, "d-mmm")).toBe("1-Jan");
  });

  it("uses the cell format in the mail table", () => {
    const book = {
      styles: {
        depth: { n: { pattern: '######.00"m"' } },
        money: { n: { pattern: '"$"#,##0.00' } },
        motor: { n: { pattern: '####.0"°/30m"' } },
        hours: { n: { pattern: '#####.00"Hrs";\\-#####.00"Hrs";' } },
      },
      sheets: {
        Data: {
          name: "Data",
          cellData: {
            "3": { "2": { v: 3186, s: "depth" } },
            "2": { "4": { v: 28672.5, s: "money" } },
            "18": { "4": { v: 4.9, s: "motor" } },
            "23": { "4": { v: 25.75, s: "hours" }, "5": { v: 0, s: "hours" } },
          },
        },
      },
    } as SnapshotBook;
    const html = rangeToHtml(book, false);
    expect(html).toContain("3186.00m");
    expect(html).toContain("$28,672.50");
    expect(html).toContain("4.9°/30m");
    expect(html).toContain("25.75Hrs");
    expect(html).not.toContain(">0<");
  });
});
