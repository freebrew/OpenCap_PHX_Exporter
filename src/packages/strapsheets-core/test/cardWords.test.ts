import { createHash } from "node:crypto";
import { inflateSync } from "node:zlib";
import { describe, expect, it } from "vitest";
import { zlibDeflate } from "../src/deflate";
import { renderCards, type CardOptions, type EntryPayload, type TextImage, type TextImager } from "../src/index";

const MOTOR: EntryPayload = {
  which: "Motor",
  cells: {
    2: "PHX-725047", 3: "Steel", 4: "BOX", 5: "PIN", 6: "4 1/2 IF [NC 50]", 7: "4 1/2 REG",
    8: "0.23", 9: "0.56", 10: "0.71", 11: "1.37", 12: "2.03", 13: "2.02", 14: "9.4",
    15: "10.17", 16: "172", 17: "184", 18: "205", 19: "184", 20: "178", 21: "169", 22: "168", 23: "1.83",
    186: "177", bendHousing: "ADJ",
  },
};

const FLOAT: EntryPayload = {
  which: "Subs",
  cells: { 90: "FLT-9", 91: "Steel", 92: "BOX", 93: "PIN", 94: "4 1/2 IF [NC 50]", 95: "4 1/2 REG", 96: "Float", 97: "0.69", 98: "169", 99: "169", 100: "75" },
};

const CUSTOM: EntryPayload = {
  which: "Custom",
  cells: {
    serial: "CUST-77",
    material: "Steel",
    boxUpper: "BOX",
    boxLower: "PIN",
    threadUpper: "4 1/2 IF",
    threadLower: "3 1/2 REG",
    totalLength: "11.40",
    internalId: "2.25",
    marksJson: JSON.stringify({
      marks: [
        { letter: "A", kind: "upset", x: 0.28, value: "6.75" },
        { letter: "B", kind: "bentHousing", x: 0.5, value: "7.10", length: "4.20", housing: "ADJ", setting: "1.50" },
        { letter: "C", kind: "length", x: 0.78, value: "3.10" },
      ],
    }),
  },
};

const MWD: EntryPayload = { which: "MWD Carrier", cells: { serial: "MWD-5", xoverSn: "SN1" }, at: "2026-01-02T12:00:00.000Z" };

const latin = (pdf: Uint8Array) => new TextDecoder("latin1").decode(pdf);
const strings = (pdf: Uint8Array) => [...latin(pdf).matchAll(/\(((?:[^()\\]|\\.)*)\) Tj/g)].map((match) => match[1] ?? "");
const sha = (pdf: Uint8Array) => createHash("sha256").update(pdf).digest("hex");

/** A stand-in for the phone's canvas: 4 pixels a point, a solid bar where the ink would be. */
function stubImager(calls: { text: string; rtl: boolean }[] = []): TextImager {
  return (text, sizePt, bold, rtl) => {
    calls.push({ text, rtl });
    const widthPt = Array.from(text).length * sizePt * (bold ? 0.62 : 0.56);
    const heightPt = sizePt * 1.25;
    const pixelWidth = Math.max(1, Math.ceil(widthPt * 4));
    const pixelHeight = Math.max(1, Math.ceil(heightPt * 4));
    const alpha = new Uint8Array(pixelWidth * pixelHeight);
    for (let y = Math.floor(pixelHeight * 0.2); y < Math.floor(pixelHeight * 0.8); y++) alpha.fill(255, y * pixelWidth, (y + 1) * pixelWidth);
    const image: TextImage = { widthPt, heightPt, baselinePt: sizePt, pixelWidth, pixelHeight, alpha };
    return image;
  };
}

const table = (words: Record<string, string>) => (english: string) => words[english] ?? english;

describe("English cards stay exactly as they were", () => {
  // Hashes of the cards as the core draws them in English (with the torque reference strip). A deliberate card change updates them.
  const BEFORE: Record<string, string> = {
    motor: "533a8fbff474adba266fd2dac5d621557ad6104bbcdb5c67889fb7983f50d463",
    float: "af4d78870f3d40c7a1731b286eb1487634f7175a4e9fe3c237ef03817a24c7b6",
    custom: "d78170814ddf0d15bc011097dd520278f2d46dc355ea6103f8190d30c9841873",
    mwd: "cd3669a9440c24527d10a5f81315c7309f2e5fe94d8a1a955eb9a29e7da87b84",
    all: "df0505b5248798f4ce09f782ed05a3503a90de533e587db6f0adb99ceb89d322",
  };
  const cases: Record<string, EntryPayload[]> = { motor: [MOTOR], float: [FLOAT], custom: [CUSTOM], mwd: [MWD], all: [MOTOR, FLOAT, CUSTOM, MWD] };

  it("draws the same bytes with no options, with an identity word, and with an imager that has nothing to draw", () => {
    for (const [name, entries] of Object.entries(cases)) {
      const plain = renderCards(entries);
      expect(sha(plain), name).toBe(BEFORE[name]);
      expect(sha(renderCards(entries, {})), name).toBe(BEFORE[name]);
      expect(sha(renderCards(entries, { word: (english) => english })), name).toBe(BEFORE[name]);
      const calls: { text: string; rtl: boolean }[] = [];
      expect(sha(renderCards(entries, { word: (english) => english, textImage: stubImager(calls) })), name).toBe(BEFORE[name]);
      expect(calls, name).toEqual([]);
    }
  });
});

describe("a card in another language", () => {
  const SPANISH = table({
    Serial: "SERIE",
    Material: "MATERIAL X",
    Steel: "ACERO",
    BOX: "CAJA",
    PIN: "PERNO",
    "4 1/2 IF [NC 50]": "4 1/2 IFX [NC 50]",
    MOTORS: "MOTORES",
    LEN: "LONG",
    OD: "DE",
    ADJ: "AJU",
    CUSTOM: "PERSONALIZADO",
    "FLOAT Subs": "SUBS FLOTADOR",
  });

  it("puts every label, title and saved name through word, and keeps serials and numbers as typed", () => {
    const asked: string[] = [];
    const word = (english: string) => {
      asked.push(english);
      return SPANISH(english);
    };
    const drawn = strings(renderCards([MOTOR, FLOAT, CUSTOM], { word }));
    for (const shown of ["SERIE", "MATERIAL X", "ACERO", "MOTORES", "SUBS FLOTADOR", "PERSONALIZADO", "H LONG", "I DE", "A DE", "B AJU", "PERNO", "CAJA"]) {
      expect(drawn, shown).toContain(shown);
    }
    expect(drawn.join("|")).toMatch(/CAJA 4 1\/2\|IFX \[NC 50\]/);
    for (const kept of ["PHX-725047", "FLT-9", "CUST-77", "0.23", "10.17", "11.40"]) expect(drawn, kept).toContain(kept);
    for (const english of ["Serial", "Steel", "MOTORS", "CUSTOM"]) expect(drawn, english).not.toContain(english);
    for (const typed of ["PHX-725047", "FLT-9", "CUST-77", "0.23", "10.17", "172"]) expect(asked, typed).not.toContain(typed);
  });

  it("decides the yellow block and the value sizes from the English label", () => {
    const english = latin(renderCards([MOTOR]));
    const strange = latin(renderCards([MOTOR], { word: table({ "H LEN": "ZZ", "I OD": "QQ", Serial: "NUMERO", TOP: "ARRIBA" }) }));
    const yellow = (source: string) => source.split("1 0.95 0.8 rg").length;
    expect(yellow(strange)).toBe(yellow(english));
    const sizeOf = (source: string, value: string) => Number(new RegExp(`/F2 (\\d+(?:\\.\\d+)?) Tf [^\\n]*\\(${value}\\) Tj`).exec(source)?.[1]);
    expect(sizeOf(strange, "PHX-725047")).toBe(sizeOf(english, "PHX-725047"));
    expect(sizeOf(strange, "0.23")).toBe(sizeOf(english, "0.23"));
    expect(sizeOf(strange, "0.23")).toBeGreaterThan(sizeOf(strange, "PHX-725047"));
  });

  it("prints the fixed notices and the unknown-tool list in the card's language", () => {
    expect(strings(renderCards([], { word: table({ "No strap sheets chosen.": "NADA ELEGIDO." }) }))).toContain("NADA ELEGIDO.");
    const odd = strings(renderCards([{ which: "Mystery", cells: { 4: "Steel" } }], { word: table({ "Column {n}": "Columna {n}", Steel: "ACERO" }) }));
    expect(odd).toContain("Columna 4");
    expect(odd).toContain("ACERO");
  });
});

describe("writing Helvetica cannot draw", () => {
  const RUSSIAN = table({ Steel: "Сталь", MOTORS: "ДВИГАТЕЛИ", Serial: "Серийный №", BOX: "МУФТА", CUSTOM: "ОСОБЫЙ", LEN: "ДЛ" });

  it("draws it as a picture with a soft mask, sized by the imager, and never as question marks", () => {
    const calls: { text: string; rtl: boolean }[] = [];
    const pdf = renderCards([MOTOR, CUSTOM], { word: RUSSIAN, textImage: stubImager(calls) });
    const source = latin(pdf);
    const pictures = [...source.matchAll(/\/Tx(\d+) Do/g)].map((match) => match[1]);
    expect(pictures.length).toBeGreaterThan(0);
    expect(source).toContain("/SMask");
    expect(source.split("/ColorSpace /DeviceGray").length - 1).toBe(new Set(pictures).size);
    for (const line of strings(pdf)) expect(line, line).not.toContain("?");
    expect(calls.map((call) => call.text)).toEqual(expect.arrayContaining(["Сталь", "ДВИГАТЕЛИ", "Серийный №", "ОСОБЫЙ", "C ДЛ"]));
    expect(calls.every((call) => !call.rtl)).toBe(true);
    expect(strings(pdf)).toContain("PHX-725047");
  });

  it("centres a pictured title where the English one is centred", () => {
    const source = latin(renderCards([MOTOR], { word: RUSSIAN, textImage: stubImager() }));
    const placed = /q (\d+(?:\.\d+)?) 0 0 (\d+(?:\.\d+)?) (\d+(?:\.\d+)?) (\d+(?:\.\d+)?) cm \/Tx\d+ Do Q/g;
    const title = [...source.matchAll(placed)].map((m) => ({ w: Number(m[1]), x: Number(m[3]), y: Number(m[4]) })).sort((a, b) => b.w - a.w)[0];
    expect(title).toBeDefined();
    expect(Math.abs((title?.x ?? 0) + (title?.w ?? 0) / 2 - 306)).toBeLessThan(0.5);
    expect(title?.y ?? 99).toBeLessThan(40);
  });

  it("keeps the soft mask's cover bytes exactly", () => {
    const imager = stubImager();
    const pdf = renderCards([], { word: table({ "No strap sheets chosen.": "Ничего не выбрано." }), textImage: imager });
    const expected = imager("Ничего не выбрано.", 15, true, false).alpha;
    const bytes = Buffer.from(pdf);
    const header = latin(pdf).indexOf("/ColorSpace /DeviceGray");
    const start = latin(pdf).indexOf("stream\n", header) + "stream\n".length;
    const length = Number(/\/Length (\d+)/.exec(latin(pdf).slice(header))?.[1]);
    expect(new Uint8Array(inflateSync(bytes.subarray(start, start + length)))).toEqual(expected);
  });

  it("hands the card's direction to the imager", () => {
    const calls: { text: string; rtl: boolean }[] = [];
    const options: CardOptions = { word: table({ MOTORS: "محركات" }), textImage: stubImager(calls), rtl: true };
    renderCards([MOTOR], options);
    expect(calls.find((call) => call.text === "محركات")?.rtl).toBe(true);
  });

  it("keeps a typed thread left to right inside right-to-left words, and plain typed text in Helvetica", () => {
    const calls: { text: string; rtl: boolean }[] = [];
    const pdf = renderCards([MOTOR], { word: table({ BOX: "صندوق", MOTORS: "محركات" }), textImage: stubImager(calls), rtl: true });
    expect(calls.some((call) => call.text.startsWith("صندوق \u20664 1/2"))).toBe(true);
    expect(calls.every((call) => /[^\x20-\x7e\u2066\u2069]/.test(call.text))).toBe(true);
    const drawn = strings(pdf);
    for (const typed of ["PHX-725047", "0.23", "Steel", "PIN 4 1/2"]) expect(drawn, typed).toContain(typed);
    for (const line of drawn) expect(line, line).not.toContain("?");
  });

  it("falls back to question marks when no imager is given, as before", () => {
    const pdf = renderCards([MOTOR], { word: RUSSIAN });
    expect(strings(pdf)).toContain("?????");
    expect(latin(pdf)).not.toContain("/SMask");
  });
});

describe("zlibDeflate", () => {
  it("makes streams zlib reads back, and shrinks runs", () => {
    const runs = new Uint8Array(200_000);
    for (let index = 0; index < runs.length; index++) runs[index] = Math.floor(index / 997) % 3 === 0 ? 255 : 0;
    let seed = 7;
    const noise = Uint8Array.from({ length: 50_000 }, () => (seed = (seed * 1103515245 + 12345) & 0x7fffffff) & 255);
    for (const data of [new Uint8Array(0), Uint8Array.of(1), Uint8Array.of(1, 2, 3, 1, 2, 3, 1, 2, 3), runs, noise]) {
      expect(new Uint8Array(inflateSync(zlibDeflate(data)))).toEqual(data);
    }
    expect(zlibDeflate(runs).length).toBeLessThan(runs.length / 50);
  });
});
