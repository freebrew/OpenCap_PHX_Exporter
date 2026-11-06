import { describe, expect, it } from "vitest";
import { PRINT_CARDS } from "../src/cards.gen";
import {
  ENTRY_FAMILIES,
  ENTRY_WHICH_OPTIONS,
  entryLines,
  findFamily,
  liveValue,
  mwdCardEntry,
  mwdRounds,
  mwdSubLengths,
  mwdSummary,
  normalizeEntry,
  renderCards,
  setCell,
  stepColumns,
  stepInput,
  stepProblem,
  validateEntry,
  type Cells,
  type EntryFamily,
  type EntryPayload,
  type RoundStep,
} from "../src/index";

const family = (): EntryFamily => {
  const found = findFamily(ENTRY_FAMILIES, "MWD Carrier");
  if (!found) throw new Error("no MWD Carrier family");
  return found;
};
const round = (id: string): RoundStep => {
  const step = family().steps.find((candidate) => candidate.id === id);
  if (!step || step.kind !== "round") throw new Error(`no ${id} round`);
  return step;
};

/** A carrier typed the way the wizard stores it: lengths are running totals from the pin. */
function typed(): Cells {
  let cells: Cells = { serial: "MWD-7", boxUpper: "BOX", boxLower: "PIN", threadUpper: "4 1/2 IF", threadLower: "4 1/2 REG", id: "70.50" };
  ["sn1", "sn2", "sn3", "sn4"].forEach((value, index) => {
    cells = setCell(cells, round("sn").cols[index]?.col ?? "", value.toUpperCase());
  });
  ["1.2", "3.5", "6", "7.8"].forEach((value, index) => {
    cells = setCell(cells, round("length").cols[index]?.col ?? "", value);
  });
  ["120.50", "121", "118", "119.25"].forEach((value, index) => {
    cells = setCell(cells, round("od").cols[index]?.col ?? "", value);
  });
  return cells;
}

const text = (pdf: Uint8Array) => new TextDecoder("latin1").decode(pdf);
const strings = (pdf: Uint8Array) => [...text(pdf).matchAll(/\(((?:[^()\\]|\\.)*)\) Tj/g)].map((match) => match[1]);

describe("the MWD Carrier on the phone", () => {
  it("is offered, with the sheet's steps: carrier, ends, three rounds, then one ID", () => {
    expect(ENTRY_WHICH_OPTIONS).toContain("MWD Carrier");
    expect(family().mwd).toBe(true);
    expect(family().steps.map((step) => `${step.kind}:${step.id}`)).toEqual([
      "text:serial", "boxPin:boxPin", "threads:threads", "round:sn", "round:length", "round:od", "decimal:id",
    ]);
    expect(round("sn")).toMatchObject({ decimal: false, cumulative: false });
    expect(round("length")).toMatchObject({ decimal: true, cumulative: true, unit: "m" });
    expect(round("od")).toMatchObject({ decimal: true, cumulative: false, unit: "mm" });
    expect(round("length").cols.map((sub) => `${sub.letter} ${sub.short}`)).toEqual(["A X-Over", "B Grounding", "C Centralizer", "D Gap"]);
    expect(stepColumns(round("od"))).toEqual(["xoverOd", "groundingOd", "centralizerOd", "gapOd"]);
    expect(stepInput(round("sn"))).toBe("round");
    expect(mwdRounds(family())).not.toBeNull();
  });

  it("types serials in capitals and measures as numbers", () => {
    expect(liveValue(round("sn"), "ab-12")).toBe("AB-12");
    expect(liveValue(round("length"), "1,2.5.3x")).toBe("12.53");
  });

  it("lets a blank round through but stops a bad number or a total that goes backwards", () => {
    expect(stepProblem(round("length"), {})).toBeNull();
    expect(stepProblem(round("od"), { xoverOd: "abc" })).toMatch(/A X-Over must be a number/);
    expect(stepProblem(round("length"), { xoverLength: "2", groundingLength: "1.5" })).toBe(
      "B Grounding total (1.5) is less than the previous total (2).",
    );
    expect(stepProblem(round("length"), { xoverLength: "2", centralizerLength: "2.5" })).toBeNull();
    expect(stepProblem(round("length"), { xoverLength: "2", groundingLength: "2" })).toBeNull();
  });

  it("stores each sub's own length and works out the total and largest OD as the sheet does", () => {
    expect(mwdSubLengths(["1.2", "3.5", "6", "7.8"])).toEqual(["1.2", "2.3", "2.5", "1.8"]);
    expect(mwdSubLengths(["1.2", "", "6", ""])).toEqual(["1.2", "", "4.8", ""]);
    expect(mwdSummary(family(), typed())).toEqual({ subLengths: ["1.2", "2.3", "2.5", "1.8"], total: "7.8", maxOd: "121" });
    expect(mwdSummary(family(), {})).toEqual({ subLengths: ["", "", "", ""], total: "", maxOd: "" });
  });

  it("validates and normalises a whole carrier", () => {
    const entry = normalizeEntry(family(), { which: "MWD Carrier", cells: typed() });
    expect(entry.cells.serial).toBe("MWD-7");
    expect(entry.cells.xoverSn).toBe("SN1");
    expect(validateEntry(family(), entry)).toEqual([]);
    const empty = validateEntry(family(), normalizeEntry(family(), { which: "MWD Carrier", cells: {} }));
    expect(empty.map((problem) => problem.stepId)).toEqual(["serial", "boxPin", "threads"]);
    const backwards = validateEntry(family(), { which: "MWD Carrier", cells: { ...typed(), groundingLength: "0.5" } });
    expect(backwards.some((problem) => problem.stepId === "length")).toBe(true);
    expect(entryLines(family(), typed()).map((line) => line.label)).toContain("B Grounding Length");
  });

  it("puts every value of the sheet's MWD.P card under its label", () => {
    const entry = mwdCardEntry(family(), { which: "MWD Carrier", cells: normalizeEntry(family(), { which: "MWD Carrier", cells: typed() }).cells, at: new Date(2026, 10, 6, 15, 30).toISOString() });
    const card = PRINT_CARDS.find((candidate) => candidate.which === "MWD Carrier");
    if (!card) throw new Error("no MWD card");
    const expected: Record<string, string> = {
      Carrier: "MWD-7", Date: "2026-11-06",
      "A SN": "SN1", "B SN": "SN2", "C SN": "SN3", "D SN": "SN4",
      "TOT LEN": "7.8", "MAX OD": "121", ID: "70.5",
      "A LEN": "1.2", "B LEN": "2.3", "C LEN": "2.5", "D LEN": "1.8",
      "A OD": "120.5", "B OD": "121", "C OD": "118", "D OD": "119.25",
      TOP: "BOX 4 1/2 IF", BTM: "PIN 4 1/2 REG",
    };
    const labelAbove = (row: number, column: number) => card.statics.find(([r, c]) => r === row - 1 && c === column)?.[2] ?? "";
    for (const [row, column, from, also] of card.fields) {
      const label = labelAbove(row, column);
      const value = [from, also].map((source) => (source == null ? "" : (entry.cells[source] ?? ""))).filter(Boolean).join(" ");
      expect(value, `${label} at row ${row} column ${column}`).toBe(expected[label]);
    }
    expect(card.fields.length).toBe(Object.keys(expected).length);
    expect(card.headerSpans).toEqual([[1, 2], [7, 2], [9, 2]]);
  });

  it("prints the MWD card with the sheet's labels, the drawing and the MWD title", () => {
    const entries: EntryPayload[] = [
      { which: "MWD Carrier", cells: normalizeEntry(family(), { which: "MWD Carrier", cells: typed() }).cells, at: new Date(2026, 10, 6, 15, 30).toISOString() },
    ];
    const pdf = renderCards(entries);
    const drawn = strings(pdf);
    for (const label of ["Carrier", "Date", "A SN", "B SN", "C SN", "D SN", "TOP", "BTM", "TOT LEN", "MAX OD", "ID", "A LEN", "D LEN", "A OD", "D OD"]) {
      expect(drawn, label).toContain(label);
    }
    for (const value of ["MWD-7", "2026-11-06", "SN1", "SN4", "7.8", "121", "70.5", "1.2", "2.5", "120.5", "119.25", "MWD"]) {
      expect(drawn, value).toContain(value);
    }
    expect(drawn.join("|")).toMatch(/BOX 4 1\/2 IF/);
    expect(drawn.join("|")).toMatch(/PIN 4 1\/2 REG/);
    expect(Number(/\/Count (\d+)/.exec(text(pdf))?.[1])).toBe(1);
  });
});
