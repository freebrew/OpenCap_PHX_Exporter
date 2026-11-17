import { describe, expect, it } from "vitest";
import { morningReport, morningSummaryMini } from "../../reports/src/index";

const FILLED = {
  "00:00 Depth": 4336,
  "24:00 Depth": 5432,
  "Meters Drilled": 1096,
  "Sliding Percentage": 0.32,
  "Rotating Percentage": "68%",
  "Motor S/N": "PHX-525074-PTS",
  Configuration: "1.50 Fixed",
  "Rev/L": 0.17,
  "Pump Rate": 1.25,
  Rotary: 40,
  "Today Bit Hours": 14.5,
  "Total Bit Hours": 120,
  "Today Circ Hours": "1hr8s",
  "Total Circ Hours": 130,
  "Total Motor Hours": 88,
  E13: 4.15,
  F13: "LEFT",
  E14: -2.25,
  F14: "below",
  E15: 4.72,
  "Average Motor Output": 8.5,
  "Dogleg Needed": 6.2,
  "Down Hole Temp": 92,
  "TVD @ Midnight": 2410.5,
};

describe("morningReport (Module21 GenerateMorningReport)", () => {
  it("builds the daily summary with the VBA's formatting", () => {
    const { text } = morningReport({ wellNumber: "35783", reportDate: "November 03, 2026", cells: FILLED });
    expect(text.split("\r\n").slice(0, 5)).toEqual([
      "Daily Drilling Summary \u2013 35783 November 03, 2026",
      "Drill  4336m\u2013 5432m",
      "Next 24 hrs. \u2013 ",
      "Currently \u2013 Drilling ahead no issues",
      "No issues to report.",
    ]);
    expect(text).toContain("Sliding Percentage: 32.00%");
    expect(text).toContain("Rotating Percentage: 68.00%");
    expect(text).toContain("Rev/L: 0.170");
    expect(text).toContain("Pump Rate: 1.25m\u00b3/min");
    // Motor RPM = pump m3/min x 1000 L/m3 x rev/L, then plus the rotary.
    expect(text).toContain("Motor Revs: 213 rpm");
    expect(text).toContain("Total RPM: 253 rpm");
    expect(text).toContain("Today Bit Hours: 14.5hrs.");
    expect(text).toContain("Total Bit Hours: 120hrs.");
    // More than one number in a hand-typed hours value is kept, with a trailing dot.
    expect(text).toContain("Today Circ Hours: 1hr8s.");
    expect(text).toContain("4.15m Left of P1");
    expect(text).toContain("-2.25m Below of P1");
    expect(text).toContain("Motor Output: 8.50\u00b0/30m");
    expect(text).toContain("Dogleg Needed: 6 \u00b0/30m");
    expect(text).toContain("Down Hole Temp 92\u00b0C");
    expect(text).toContain("TVD @ Midnight: 2410.5mTVD");
  });

  it("derives the start depth and reads a direction out of the distance text", () => {
    const { text } = morningReport({
      cells: { "24:00 Depth": "5432m", "Meters Drilled": 1096, E13: "4.15m LEFT", E14: "", "Above/Below from Plan P1": 2, "Above/Below from Plan P1 dir": "ABOVE" },
    });
    expect(text).toContain("00:00 Depth: 4336m");
    expect(text).toContain("4.15m Left of P1");
    expect(text).toContain("2.00m Above of P1");
  });

  it("escapes the HTML body", () => {
    const { html } = morningReport({ wellNumber: "<35783>", cells: {} });
    expect(html).toContain("&lt;35783&gt;");
    expect(html).not.toContain("<35783>");
  });
});

describe("morningSummaryMini (Module21 GenerateMorningSummaryMini)", () => {
  it("summarises the last 24 h from the Data cells", () => {
    const { text } = morningSummaryMini({
      cells: { H3: 2, I23: 4336, J23: 5432, K23: 1096, L23: 350, P23: 14.5, P24: 120, Q23: 13, Q24: 80, E7: 5400.25, E8: 62.4, E9: 185.1, E13: 4.15, E14: -2.25 },
    });
    expect(text.split("\r\n")).toEqual([
      "BHA #2",
      "Last 24Hrs",
      "4336 - 5432 =  Total 1096",
      "Rotating 746  Rotate 68%",
      "Slid 350M Slid 32%",
      "Bit Hours = 14.5hrs. Total Bit Hours = 120hrs.",
      "Motor Hours = 13hrs. Total Motor Hours = 80hrs.",
      "Last Survey = 5400.25 Inc = 62.40, Az = 185.10",
      "Well Position @ Survey = 2.25m South, 4.15m East",
    ]);
  });

  it("marks a blank cell with its address and defaults the BHA", () => {
    const { text } = morningSummaryMini({ cells: { J23: 5432, I23: 4336 } });
    expect(text.startsWith("BHA #1")).toBe(true);
    expect(text).toContain("Bit Hours = [P23]");
    expect(text).toContain("[E14]m North, [E13]m East");
    expect(text).toContain("Last Survey = [E7]");
  });
});
