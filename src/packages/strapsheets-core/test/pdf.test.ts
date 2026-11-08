import { describe, expect, it } from "vitest";
import { CARD_ART } from "../src/art.gen";
import { PRINT_CARDS } from "../src/cards.gen";
import { ENTRY_FAMILIES, renderCards, serialStep, stepColumns, type EntryPayload } from "../src/index";
import { textWidth, wrapText } from "../src/helvetica";

const MOTOR: EntryPayload = {
  which: "Motor",
  cells: {
    2: "PHX-725047", 3: "Steel", 4: "BOX", 5: "PIN", 6: "4 1/2 IF [NC 50]", 7: "4 1/2 REG",
    8: "0.23", 9: "0.56", 10: "0.71", 11: "1.37", 12: "2.03", 13: "2.02", 14: "9.4",
    15: "10.17", 16: "172", 17: "184", 18: "205", 19: "184", 20: "178", 21: "169", 22: "168", 23: "1.83",
    186: "177", bendHousing: "ADJ",
  },
};

/** A Subs entry of the given type, the columns Floats / GAPs / X-Overs read. */
function subs(serial: string, type: string): EntryPayload {
  return { which: "Subs", cells: { 90: serial, 91: "Steel", 92: "BOX", 93: "PIN", 94: "4 1/2 IF [NC 50]", 95: "4 1/2 REG", 96: type, 97: "0.69", 98: "169", 99: "169", 100: "75" } };
}

const text = (pdf: Uint8Array) => new TextDecoder("latin1").decode(pdf);
const pageCount = (pdf: Uint8Array) => Number(/\/Count (\d+)/.exec(text(pdf))?.[1]);
const strings = (pdf: Uint8Array) => [...text(pdf).matchAll(/\(((?:[^()\\]|\\.)*)\) Tj/g)].map((match) => match[1]);
const images = (pdf: Uint8Array) => text(pdf).split("/Subtype /Image").length - 1;

describe("the strap card layout", () => {
  it("draws the Sheets card: serial, connections and the lettered measures with the family title", () => {
    const drawn = strings(renderCards([MOTOR]));
    for (const label of ["Serial", "Material", "TOP", "BTM", "A", "B", "C", "D", "E", "F", "G", "BEND", "H LEN", "I OD", "J OD", "K OD", "L OD", "M OD", "N OD", "O OD", "P OD"]) {
      expect(drawn, label).toContain(label);
    }
    for (const value of ["PHX-725047", "Steel", "0.23", "9.4", "1.83", "10.17", "172", "177", "168"]) expect(drawn, value).toContain(value);
    expect(drawn).toContain("MOTORS");
  });

  it("joins the BOX/PIN end and its thread into the TOP and BTM cells, wrapped like the sheet", () => {
    const drawn = strings(renderCards([MOTOR])).join("|");
    expect(drawn).toMatch(/BOX 4 1\/2\|IF \[NC 50\]/);
    expect(drawn).toMatch(/PIN 4 1\/2\|REG/);
  });

  it("is not the old banner and list", () => {
    const drawn = strings(renderCards([MOTOR]));
    expect(drawn.some((line) => line.startsWith("Strap Sheet"))).toBe(false);
    expect(drawn).not.toContain("A - BitBox");
  });

  it("puts the tool drawing on the page as vectors", () => {
    const pdf = renderCards([MOTOR]);
    const source = text(pdf);
    expect(images(pdf)).toBe(0);
    expect(source).not.toContain("/Predictor 15");
    expect(source).toContain(" m ");
    expect(source).toContain(" l ");
    expect(source).toContain("0.14 0.19 0.22 RG");
  });

  it("puts every entry of one family on one page, one card each", () => {
    const pdf = renderCards([MOTOR, { ...MOTOR, cells: { ...MOTOR.cells, 2: "PHX-650076-PTS" } }]);
    expect(pageCount(pdf)).toBe(1);
    expect(strings(pdf)).toContain("PHX-725047");
    expect(strings(pdf)).toContain("PHX-650076-PTS");
    expect(strings(pdf).filter((line) => line === "Serial")).toHaveLength(2);
  });

  it("starts each family on its own page, in the order the sheet prints families", () => {
    const jars: EntryPayload = { which: "Jars", cells: { 115: "JAR-1" } };
    const pdf = renderCards([jars, MOTOR]);
    expect(pageCount(pdf)).toBe(2);
    const drawn = strings(pdf);
    expect(drawn.indexOf("MOTORS")).toBeLessThan(drawn.indexOf("JARS"));
  });

  it("splits Subs across the Float, GAP and X-Over layouts by sub type", () => {
    expect(strings(renderCards([subs("FLT-1", "Float")]))).toContain("FLOAT Subs");
    expect(strings(renderCards([subs("GAP-1", "Pulser")]))).toContain("GAP/PULSER SUBS");
    expect(strings(renderCards([subs("XO-1", "X-Over")]))).toContain("X-Overs");
    const mixed = renderCards([subs("FLT-1", "Float"), subs("XO-1", "X-Over"), subs("FLT-2", "Screen")]);
    expect(pageCount(mixed)).toBe(2);
  });

  it("keeps a Subs entry with no sub type instead of dropping it", () => {
    const pdf = renderCards([subs("NOTYPE-1", "")]);
    expect(strings(pdf)).toContain("NOTYPE-1");
    expect(strings(pdf)).toContain("SUBS");
  });

  it("puts two short sub cards side by side in one band", () => {
    const pdf = renderCards([subs("FLT-1", "Float"), subs("FLT-2", "Float")]);
    expect(pageCount(pdf)).toBe(1);
    const drawn = strings(pdf);
    expect(drawn).toContain("FLT-1");
    expect(drawn).toContain("FLT-2");
    expect(drawn.filter((line) => line === "RING")).toHaveLength(2);
  });

  it("starts a new page when the cards run past the bottom and prints the family title at the bottom border of each sheet", () => {
    const many = Array.from({ length: 30 }, (_, index) => ({ ...MOTOR, cells: { ...MOTOR.cells, 2: `PHX-${index}` } }));
    const pdf = renderCards(many);
    expect(pageCount(pdf)).toBeGreaterThan(1);
    expect(strings(pdf).filter((line) => line === "MOTORS")).toHaveLength(pageCount(pdf));
    const ys = [...text(pdf).matchAll(/(\d+(?:\.\d+)?) (\d+(?:\.\d+)?) Td \(MOTORS\) Tj/g)].map((m) => Number(m[2]));
    expect(ys.length).toBe(pageCount(pdf));
    for (const y of ys) expect(y).toBeLessThan(40);
    for (let index = 0; index < 30; index++) expect(strings(pdf)).toContain(`PHX-${index}`);
  });

  it("renders every family of the wizard and ends it with that family's title", () => {
    for (const family of ENTRY_FAMILIES) {
      const serial = serialStep(family);
      if (serial?.kind !== "text") throw new Error("a family has its serial");
      const entry: EntryPayload = { which: family.which, cells: { [serial.col]: "SER-1" } };
      for (const step of family.steps) if (step.kind === "choice" && step.optionsKey === "subTypeOptions") entry.cells[step.col] = "Float";
      const pdf = renderCards([entry]);
      expect(strings(pdf), family.which).toContain("SER-1");
      expect(images(pdf), family.which).toBe(0);
      expect(text(pdf), family.which).toContain(" l ");
      expect(pageCount(pdf), family.which).toBe(1);
    }
  });

  it("keeps a valid cross-reference table", () => {
    const pdf = renderCards([MOTOR, subs("GAP-1", "GAP")]);
    const source = text(pdf);
    const xref = Number(/startxref\n(\d+)\n%%EOF$/.exec(source)?.[1]);
    expect(source.slice(xref, xref + 4)).toBe("xref");
    const rows = [...source.slice(xref).matchAll(/(\d{10}) 00000 n /g)].map((match) => Number(match[1]));
    rows.forEach((offset, index) => expect(source.slice(offset, offset + `${index + 1} 0 obj`.length)).toBe(`${index + 1} 0 obj`));
  });

  it("still makes a one-page document when nothing is chosen, and keeps an entry whose family is unknown", () => {
    expect(pageCount(renderCards([]))).toBe(1);
    const odd = renderCards([{ which: "Mystery Tool", cells: { 2: "ODD-1" } }]);
    expect(strings(odd)).toContain("ODD-1");
  });
});

describe("the generated card tables", () => {
  it("print every wizard family and have a drawing for each", () => {
    for (const family of ENTRY_FAMILIES) {
      if (family.which === "Custom") continue;
      expect(PRINT_CARDS.some((card) => card.which === family.which), family.which).toBe(true);
      expect(CARD_ART[family.which], family.which).toBeDefined();
    }
  });

  it("read only columns the wizard writes, so a regenerated schema cannot leave a blank cell", () => {
    for (const card of PRINT_CARDS) {
      const family = ENTRY_FAMILIES.find((candidate) => candidate.which === card.which);
      expect(family, card.name).toBeDefined();
      // The MWD card reads the sheet's row, which mwdCardEntry builds from the named cells (mwd.test.ts).
      if (card.which === "MWD Carrier") continue;
      const written = new Set(family?.steps.flatMap(stepColumns));
      for (const field of card.fields) {
        for (const column of field.slice(2)) expect(written.has(column), `${card.name} column ${column}`).toBe(true);
      }
    }
  });

  it("give every card a label for every value cell", () => {
    for (const card of PRINT_CARDS) {
      for (const field of card.fields) {
        expect(card.statics.some(([row, col]) => row === field[0] - 1 && col === field[1]), `${card.name} ${field[0]},${field[1]}`).toBe(true);
      }
    }
  });
});

describe("Helvetica metrics", () => {
  it("measures text in the standard font's widths", () => {
    expect(textWidth("Serial", false, 1000)).toBe(2556);
    expect(textWidth("Serial", true, 1000)).toBe(2724);
    expect(textWidth("", true, 12)).toBe(0);
  });

  it("keeps a bracketed thread size together instead of splitting it across lines", () => {
    const room = textWidth("BOX 4 1/2 IF [NC", true, 14.67) + 4;
    expect(wrapText("BOX 4 1/2 IF [NC 50]", true, 14.67, room)).toEqual(["BOX 4 1/2 IF", "[NC 50]"]);
    expect(wrapText("[NC 50]", true, 14.67, textWidth("[NC", true, 14.67) + 2).join(" ")).toBe("[NC 50]");
  });

  it("wraps at spaces, then inside a word that is still too wide", () => {
    expect(wrapText("BOX 4 1/2 IF [NC 50]", true, 14.67, 74)).toEqual(["BOX 4 1/2", "IF [NC 50]"]);
    expect(wrapText("SHORT", true, 14.67, 74)).toEqual(["SHORT"]);
    const broken = wrapText("ABCDEFGHIJKLMNOPQRSTUVWXYZ", true, 14.67, 74);
    expect(broken.length).toBeGreaterThan(1);
    expect(broken.join("")).toBe("ABCDEFGHIJKLMNOPQRSTUVWXYZ");
  });
});
