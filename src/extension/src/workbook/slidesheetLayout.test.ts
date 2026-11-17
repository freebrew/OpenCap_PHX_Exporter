import { describe, expect, it } from "vitest";
import { normalizeSlidesheetRowHeights } from "./slidesheetLayout";

describe("Slidesheet layout", () => {
  it("uses the survey table's modal explicit height for every survey row", () => {
    const rowData: Record<number, { h: number }> = {
      11: { h: 28 },
      12: { h: 19 },
      13: { h: 20 },
      14: { h: 20 },
      15: { h: 20 },
      16: { h: 20 },
      17: { h: 20 },
      18: { h: 20 },
      19: { h: 20 },
      20: { h: 20 },
      123: { h: 19 },
      140: { h: 19 },
      185: { h: 19 },
      505: { h: 30 },
    };
    const snapshot = {
      sheets: {
        slides: {
          name: "Slidesheet",
          defaultRowHeight: 19,
          rowData,
        },
      },
    };

    normalizeSlidesheetRowHeights(snapshot);

    expect(snapshot.sheets.slides.rowData[11]?.h).toBe(28);
    expect(snapshot.sheets.slides.rowData[12]?.h).toBe(20);
    expect(snapshot.sheets.slides.rowData[123]?.h).toBe(20);
    expect(snapshot.sheets.slides.rowData[140]?.h).toBe(20);
    expect(snapshot.sheets.slides.rowData[185]?.h).toBe(20);
    expect(snapshot.sheets.slides.rowData[504]?.h).toBe(20);
    expect(snapshot.sheets.slides.rowData[505]?.h).toBe(30);
  });
});
