import { describe, expect, it } from "vitest";
import { ENTRY_FAMILIES, ENTRY_WHICH_OPTIONS, entryLines, findFamily, liveValue, measureSteps, normalizeEntry, prefillFromItem, sealApplies, sealParts, sealSide, sealText, sealValue, serialStep, setCell, setRange, stepColumns, stepProblem, stepWarning, threadParts, validateEntry, type Cells, type EntryFamily, type EntryStep, type InventoryItem } from "../src/index";

function motor(): EntryFamily {
  const family = findFamily(ENTRY_FAMILIES, "Motor");
  if (!family) throw new Error("Motor family missing");
  return family;
}

function step(family: EntryFamily, id: string): EntryStep {
  const found = family.steps.find((candidate) => candidate.id === id);
  if (!found) throw new Error(`${id} missing`);
  return found;
}

const item: InventoryItem = { itemCode: "M1", itemName: "Motor 6.75", serialNumber: "abc123", mnfNumber: "", category: "Motor", subCategory: "", status: "On Job", runHours: 12 };

describe("thread names", () => {
  it("splits a bracketed second name and fills in the known ones", () => {
    expect(threadParts("4 1/2 IF [NC 50]")).toEqual({ name: "4 1/2 IF", alias: "NC 50" });
    expect(threadParts("DS 38")).toEqual({ name: "DS 38", alias: "NC38" });
    expect(threadParts("HLNST-43")).toEqual({ name: "HLNST-43", alias: "Hilong" });
    expect(threadParts("6 5/8 REG")).toEqual({ name: "6 5/8 REG", alias: null });
  });
});

describe("wizard", () => {
  it("offers every front-page tool and each has a family", () => {
    expect(ENTRY_WHICH_OPTIONS).toHaveLength(14);
    for (const which of ENTRY_WHICH_OPTIONS) expect(findFamily(ENTRY_FAMILIES, which)).not.toBeNull();
  });

  it("forces text to capitals and keeps decimals to digits and one dot", () => {
    const family = motor();
    expect(liveValue(step(family, "serial"), "abc-12x")).toBe("ABC-12X");
    expect(liveValue(step(family, "d6"), "1a2.3.4")).toBe("12.34");
  });

  it("stops only on a blank required step or a bad number", () => {
    const family = motor();
    const serial = step(family, "serial");
    expect(stepProblem(serial, {})).toBe("Serial is required.");
    if (serial.kind !== "text" || !serial.companionChoice) throw new Error("Motor Bend Housing missing");
    expect(stepProblem(serial, { [serial.col]: "M-1" })).toBe("Bend Housing is required.");
    expect(stepProblem(serial, { [serial.col]: "M-1", [serial.companionChoice.col]: "FXD" })).toBeNull();
    expect(stepProblem(step(family, "d6"), {})).toBeNull();
    expect(stepProblem(step(family, "d6"), { 8: "." })).toBe("A - BitBox needs a number.");
  });

  it("warns on an out-of-range measure and never blocks", () => {
    const family = motor();
    const bitBox = step(family, "d6");
    const ranges = setRange([], { family: "Motor", stepId: "d6", min: 5, max: 10 });
    expect(stepWarning(family, bitBox, { 8: "12" }, ranges)).toBe("above");
    expect(stepWarning(family, bitBox, { 8: "7" }, ranges)).toBeNull();
    expect(stepProblem(bitBox, { 8: "12" })).toBeNull();
  });

  it("prefills the picked serial and earlier measures into blanks only", () => {
    const family = motor();
    const history = [{ which: "Motor", cells: { 2: "ABC123", 8: "9.5", 9: "3" } }];
    const cells = prefillFromItem(family, { 9: "4" }, item, history);
    expect(cells[2]).toBe("ABC123");
    expect(cells[8]).toBe("9.5");
    expect(cells[9]).toBe("4");
  });

  it("completes a Motor entry typed with no snapshot", () => {
    const family = motor();
    let cells: Cells = {};
    for (const current of family.steps) {
      switch (current.kind) {
        case "text":
          cells = setCell(cells, current.col, liveValue(current, "mtr-001"));
          if (current.companionChoice) {
            cells = setCell(cells, current.companionChoice.col, "ADJ");
          }
          break;
        case "choice":
          cells = setCell(cells, current.col, family.materials[0] ?? "");
          break;
        case "boxPin":
          cells = setCell(setCell(cells, current.cols.upper, "BOX"), current.cols.lower, "PIN");
          break;
        case "threads":
          cells = setCell(setCell(cells, current.cols.upper, "4 1/2 IF [NC 50]"), current.cols.lower, "4 1/2 REG");
          break;
        case "decimal":
          if (current.id !== "d7") cells = setCell(cells, current.col, "1.5");
          break;
        case "customBuilder":
        case "round":
          break;
        default: {
          const never: never = current;
          throw new Error(String(never));
        }
      }
      expect(stepProblem(current, cells)).toBeNull();
    }
    const payload = normalizeEntry(family, { which: "Motor", cells });
    expect(validateEntry(family, payload)).toEqual([]);
    expect(payload.cells[2]).toBe("MTR-001");
    expect(payload.cells.bendHousing).toBe("ADJ");
    const skipped = stepColumns(step(family, "d7"))[0];
    expect(skipped).toBeDefined();
    expect(payload.cells[skipped ?? -1]).toBeUndefined();
    expect(entryLines(family, payload.cells).slice(0, 2)).toEqual([
      { label: "Serial", value: "MTR-001" },
      { label: "Bend Housing", value: "ADJ" },
    ]);
    expect(measureSteps(family).length).toBe(17);
  });

  it("lists a Subs confirm in column order, so Serial stays above Sub Type", () => {
    const family = findFamily(ENTRY_FAMILIES, "Subs");
    if (!family) throw new Error("Subs family missing");
    const subType = step(family, "c6");
    const serial = serialStep(family);
    if (subType.kind !== "choice" || serial?.kind !== "text") throw new Error("Subs steps");
    const cells = { [subType.col]: "Float", [serial.col]: "FLT-1" };
    expect(entryLines(family, cells).map((line) => line.label)).toEqual(["Serial", "Sub Type"]);
  });
});

describe("UBHO seal measure", () => {
  function subsFamily(): EntryFamily {
    const family = findFamily(ENTRY_FAMILIES, "Subs");
    if (!family) throw new Error("Subs family missing");
    return family;
  }
  function sealed(family: EntryFamily): EntryStep {
    const found = family.steps.find((candidate) => candidate.kind === "decimal" && candidate.seal);
    if (!found) throw new Error("Subs has no seal step");
    return found;
  }

  it("offers UBHO in place of Pulser and keeps the seal on the ring column", () => {
    const family = subsFamily();
    expect(family.extraOptions?.subTypeOptions).toContain("UBHO");
    expect(family.extraOptions?.subTypeOptions).not.toContain("Pulser");
    const ring = sealed(family);
    expect(ring.kind === "decimal" && ring.col).toBe(101);
    expect(ring.kind === "decimal" && ring.seal).toEqual({ typeCol: 96, types: ["UBHO"] });
  });

  it("shows the two-seal panel only on a UBHO", () => {
    const ring = sealed(subsFamily());
    expect(sealApplies(ring, { 96: "UBHO" })).toBe(true);
    expect(sealApplies(ring, { 96: "GAP" })).toBe(false);
    expect(sealApplies(ring, {})).toBe(false);
  });

  it("keeps a TOP number with its prefix and a BTM number as typed", () => {
    expect(sealParts("TOP 3.25")).toEqual({ btm: "", top: "3.25" });
    expect(sealParts("2.5")).toEqual({ btm: "2.5", top: "" });
    expect(sealValue({ btm: "", top: " 3.25 " })).toBe("TOP 3.25");
    expect(sealValue({ btm: "2.5", top: "" })).toBe("2.5");
    expect(sealSide(sealParts("TOP 1"))).toBe("top");
    expect(sealSide(sealParts("1"))).toBe("btm");
    expect(sealSide(sealParts(""))).toBe("");
    expect(sealText("top 3.25")).toBe("TOP 3.25");
    expect(sealText("TOP x")).toBeNull();
    expect(sealText("")).toBe("");
  });

  it("lets a TOP measure through the wizard and onto the stored entry", () => {
    const family = subsFamily();
    const ring = sealed(family);
    expect(stepProblem(ring, { 96: "UBHO", 101: "TOP 3.25" })).toBeNull();
    expect(stepProblem(ring, { 96: "UBHO", 101: "TOP" })).not.toBeNull();
    const payload = { which: "Subs", cells: { 90: "UBHO-1", 96: "UBHO", 101: "TOP 3.25" } };
    expect(validateEntry(family, payload).filter((problem) => problem.stepId === ring.id)).toEqual([]);
    expect(normalizeEntry(family, payload).cells["101"]).toBe("TOP 3.25");
  });
});