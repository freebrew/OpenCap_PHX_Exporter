import { describe, expect, it } from "vitest";
import { choiceOptions, decimalText, entryPayloadSchema, filterItems, itemsFromCsv, normalizeEntry, rangeFor, rangeWarning, setRange, stepColumns, stepInput, upperText, validateEntry, type EntryFamily, type EntryStep, type InventoryItem, ENTRY_FAMILIES } from "../src/index";

// The Motor family, trimmed to the steps the wizard opens with. The order and the input mode
// are what the snapshot below pins, so a regenerated entry_schema.js that drifts fails here.
const MOTOR: EntryFamily = {
  which: "Motor",
  materials: ["Steel", "Non Magnetic"],
  boxPinOptions: ["PIN", "BOX"],
  extraOptions: {},
  steps: [
    { id: "serial", kind: "text", label: "Serial", header: "Serial", col: 2, required: true },
    { id: "material", kind: "choice", label: "Material", header: "Material", col: 3, optionsKey: "materials", required: true },
    { id: "boxPin", kind: "boxPin", label: "BOX PIN Config", cols: { upper: 4, lower: 5 }, headers: { upper: "BOX PIN Config [Upper]", lower: "BOX PIN Config [Lower]" } },
    { id: "threads", kind: "threads", label: "BOX PIN Thread Type", cols: { upper: 6, lower: 7 }, headers: { upper: "BOX PIN Thread Type [Upper Thread Type]", lower: "BOX PIN Thread Type [Lower Thread Type]" } },
    { id: "d6", kind: "decimal", label: "A - BitBox", header: "A - BitBox", col: 8 },
    { id: "d7", kind: "decimal", label: "B - Lower PAD", header: "B - Lower PAD", col: 9 },
  ],
};

describe("the wizard renders from the schema", () => {
  it("keeps the step order and the input mode for each", () => {
    expect(MOTOR.steps.map((step) => step.id)).toEqual(["serial", "material", "boxPin", "threads", "d6", "d7"]);
    expect(MOTOR.steps.map(stepInput)).toEqual(["text", "buttons", "buttons", "buttons", "decimal", "decimal"]);
  });

  it("maps a step onto the columns it fills", () => {
    expect(MOTOR.steps.flatMap(stepColumns)).toEqual([2, 3, 4, 5, 6, 7, 8, 9]);
  });

  it("offers a choice step its list", () => {
    const material = MOTOR.steps[1];
    if (material?.kind !== "choice") throw new Error("material is a choice");
    expect(choiceOptions(MOTOR, material)).toEqual(["Steel", "Non Magnetic"]);
  });
});

describe("the generated schema", () => {
  it("opens every family with a serial typed in capitals, then buttons, then the number pad", () => {
    for (const family of ENTRY_FAMILIES) {
      const first = family.steps[0] as EntryStep;
      expect(first.kind, family.which).toBe("text");
      expect(first.id, family.which).toBe("serial");
      for (const step of family.steps) {
        if (step.kind === "decimal") expect(stepInput(step), `${family.which} ${step.id}`).toBe("decimal");
        if (step.kind === "text") expect(stepInput(step)).toBe("text");
      }
    }
    expect(ENTRY_FAMILIES.map((family) => family.which)).toEqual([
      "Motor", "NM Slick Monel", "NM Flex Monel", "NM Pony", "NM Stabilizer Carrier", "Stabilized Sub",
      "Subs", "Lifting Subs", "Jars", "Shock Sub", "Exciter / Agitator / Sidewinder", "Orbit", "MWD Carrier", "Custom",
    ]);
  });

  it("rejects a payload that is not an entry", () => {
    expect(entryPayloadSchema.safeParse({ which: "Motor", cells: { "2": "PHX-1" } }).success).toBe(true);
    expect(entryPayloadSchema.safeParse({ which: "", cells: {} }).success).toBe(false);
  });
});

describe("entry text", () => {
  it("forces capitals and accepts a decimal or a blank", () => {
    expect(upperText("  phx-525074-pts ")).toBe("PHX-525074-PTS");
    expect(decimalText("6.75")).toBe("6.75");
    expect(decimalText("")).toBe("");
    expect(decimalText("6.7.5")).toBeNull();
  });

  it("rejects a missing serial and a material that is not offered", () => {
    const problems = validateEntry(MOTOR, { which: "Motor", cells: { 3: "Titanium" } });
    expect(problems.map((problem) => problem.stepId)).toEqual(["serial", "material"]);
  });

  it("stores the serial in capitals and drops blanks", () => {
    const stored = normalizeEntry(MOTOR, { which: "Motor", cells: { 2: "phx-1", 3: "Steel", 8: "" } });
    expect(stored.cells).toEqual({ 2: "PHX-1", 3: "Steel" });
  });
});

describe("inventory", () => {
  const csv = "ItemCode,ItemName,SerialNumber,MnfNumber,Category,SubCategory,Status,HrsSinceService\r\nMOT,Mud Motor,phx-525074-pts,MF1,DD,Motor,Active,120\r\n,Bit,,MF2,DD,Bit,Active,\r\n";

  it("reads the exporter's CSV, drops a row with no serial, and keeps capitals", () => {
    const items = itemsFromCsv(csv);
    expect(items).toHaveLength(1);
    expect(items[0]).toMatchObject({ serialNumber: "PHX-525074-PTS", itemName: "Mud Motor", category: "DD", runHours: 120 });
  });

  it("filters by serial, name or category", () => {
    const items: InventoryItem[] = [
      { itemCode: "MOT", itemName: "Mud Motor", serialNumber: "PHX-1", mnfNumber: "", category: "DD", subCategory: "Motor", status: "", runHours: null },
      { itemCode: "AG", itemName: "Agitator", serialNumber: "AG-100", mnfNumber: "", category: "DD", subCategory: "", status: "", runHours: null },
    ];
    expect(filterItems(items, "agit").map((item) => item.serialNumber)).toEqual(["AG-100"]);
    expect(filterItems(items, "").length).toBe(2);
  });
});

describe("measure ranges", () => {
  it("warns below and above, and never without a range", () => {
    const range = { family: "Motor", stepId: "d6", min: 4, max: 12 };
    expect(rangeWarning(range, 3)).toBe("below");
    expect(rangeWarning(range, 14)).toBe("above");
    expect(rangeWarning(range, 8)).toBeNull();
    expect(rangeWarning(undefined, 3)).toBeNull();
  });

  it("stores a range low to high and finds it again", () => {
    const stored = setRange([], { family: "Motor", stepId: "d6", min: 12, max: 4 });
    expect(rangeFor(stored, "Motor", "d6")).toEqual({ family: "Motor", stepId: "d6", min: 4, max: 12 });
    expect(rangeFor(stored, "Motor", "d7")).toBeUndefined();
  });
});
