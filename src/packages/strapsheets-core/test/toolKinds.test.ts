import { describe, expect, it } from "vitest";
import { ENTRY_FAMILIES, findFamily, measuredSerials, serialStep, strapCandidate, type InventoryItem } from "../src/index";

const item = (serialNumber: string, itemName: string, category: string, subCategory = ""): InventoryItem => ({
  itemCode: "",
  itemName,
  serialNumber,
  mnfNumber: "",
  category,
  subCategory,
  status: "",
  runHours: null,
});

describe("strap candidates (the Data tab's 3rd Party Tools rule)", () => {
  it("keeps Other tools and drops the kinds the 3rd Party table skips", () => {
    expect(strapCandidate(item("266", "Shock Sub 6 3/4", "Other Tools"))).toBe(true);
    expect(strapCandidate(item("J-101", "Drilling Jar", "Other Inventory", "Jars"))).toBe(true);
    expect(strapCandidate(item("BIT-9", "8 3/4 PDC Drill Bit", "Other Tools"))).toBe(false);
    expect(strapCandidate(item("DP-1", "Drill Pipe 4 1/2", "Other"))).toBe(false);
    expect(strapCandidate(item("HW-1", "HWDP 5", "Other Tools"))).toBe(false);
    expect(strapCandidate(item("SC-1", "Slick Collar 6 1/2", "Other Tools"))).toBe(true);
    expect(strapCandidate(item("PC-1", "Pony Collar 6 1/2", "Other Tools", "DD Tubular"))).toBe(true);
    expect(strapCandidate(item("NM-5", "NM Slick Monel 6 3/4", "Other Inventory"))).toBe(true);
    expect(strapCandidate(item("KIT-1", "Probe", "Other Tools", "MWD Kit"))).toBe(false);
    expect(strapCandidate(item("TB-1", "Crossover", "Other Tools", "DD Tubular"))).toBe(false);
  });

  it("keeps motors in any category, and DD categories untouched", () => {
    expect(strapCandidate(item("PHX-576154-PTS", "6 3/4 Mud Motor", "Other Tools"))).toBe(true);
    expect(strapCandidate(item("ORB-7", "Orbit RSS 675", "MWD"))).toBe(true);
    expect(strapCandidate(item("NM-22", "NM Flex Monel", "DD Tools"))).toBe(true);
  });

  it("drops rows whose serial is not a serial", () => {
    expect(strapCandidate(item("12", "Shock Sub", "Other Tools"))).toBe(false);
    expect(strapCandidate(item("Serial Number", "Shock Sub", "Other Tools"))).toBe(false);
  });
});

describe("measured serials", () => {
  it("collects each entry's serial, normalised so 266 and 2-66 match the inventory", () => {
    const subs = serialStep(findFamily(ENTRY_FAMILIES, "Subs")!);
    const subsCol = subs && subs.kind === "text" ? subs.col : 0;
    const serials = measuredSerials(ENTRY_FAMILIES, [
      { which: "Subs", cells: { [subsCol]: "ab-266" } },
      { which: "Motor", cells: { 2: "PHX-576154-PTS" } },
      { which: "Motor", cells: {} },
      { which: "No such tool", cells: { 2: "X1" } },
    ]);
    expect([...serials].sort()).toEqual(["AB266", "PHX576154PTS"]);
  });
});
