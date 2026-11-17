import { describe, expect, it } from "vitest";
import { ARIAL_GLYPHS } from "../../workbook-commands/src/arialGlyphs";
import {
  alignGlyphs,
  drawnGlyphs,
  fingerprintTable,
  glyphLetters,
  importAntiCollision,
  importSurveyPlan,
  packFingerprints,
  pdfDocumentText,
  readGlyphFont,
  SheetBook,
  UNKNOWN_GLYPH,
  unpackFingerprints,
  unreadablePdfText,
  type DrawnGlyph,
  type Letters,
  type PdfDocLike,
  type PdfPageLike,
  type PdfTextItem,
  type PdfTextOps,
} from "../../workbook-commands/src/index";

// ---------------------------------------------------------------- a TrueType font in memory

type Outline = { kind: "simple"; contours: Array<Array<[number, number]>> } | { kind: "composite"; parts: Array<{ gid: number; dx: number; dy: number }> } | { kind: "empty" };

function glyphBytes(outline: Outline): number[] {
  const out: number[] = [];
  const u16 = (v: number) => out.push((v >> 8) & 0xff, v & 0xff);
  if (outline.kind === "empty") return out;
  if (outline.kind === "composite") {
    u16(0xffff);
    for (let i = 0; i < 4; i += 1) u16(0);
    outline.parts.forEach((part, i) => {
      u16(1 | 2 | (i < outline.parts.length - 1 ? 0x20 : 0));
      u16(part.gid);
      u16(part.dx & 0xffff);
      u16(part.dy & 0xffff);
    });
    return out;
  }
  const points = outline.contours.flat();
  u16(outline.contours.length);
  for (let i = 0; i < 4; i += 1) u16(0);
  let end = -1;
  for (const contour of outline.contours) u16((end += contour.length));
  u16(0);
  for (let i = 0; i < points.length; i += 1) out.push(1);
  let x = 0;
  for (const [px] of points) {
    u16((px - x) & 0xffff);
    x = px;
  }
  let y = 0;
  for (const [, py] of points) {
    u16((py - y) & 0xffff);
    y = py;
  }
  return out;
}

/** head, maxp, loca (short), glyf and, when given, a format 4 cmap with one segment per code. */
function makeFont(glyphs: Outline[], cmap: Array<[number, number]> = []): Uint8Array {
  const glyf: number[] = [];
  const offsets: number[] = [];
  for (const glyph of glyphs) {
    offsets.push(glyf.length);
    glyf.push(...glyphBytes(glyph));
    if (glyf.length % 2) glyf.push(0);
  }
  offsets.push(glyf.length);
  const be = (bytes: number[], v: number, size: 2 | 4) => {
    for (let s = size - 1; s >= 0; s -= 1) bytes.push((v >>> (s * 8)) & 0xff);
  };
  const head = new Array<number>(54).fill(0);
  const maxp: number[] = [];
  be(maxp, 0x00005000, 4);
  be(maxp, glyphs.length, 2);
  const loca: number[] = [];
  for (const offset of offsets) be(loca, offset / 2, 2);
  const tables: Array<[string, number[]]> = [["glyf", glyf], ["head", head], ["loca", loca], ["maxp", maxp]];
  if (cmap.length > 0) {
    const segments = [...cmap].sort((a, b) => a[0] - b[0]).concat([[0xffff, 0]]);
    const sub: number[] = [];
    be(sub, 4, 2);
    be(sub, 16 + segments.length * 8, 2);
    be(sub, 0, 2);
    be(sub, segments.length * 2, 2);
    for (let i = 0; i < 3; i += 1) be(sub, 0, 2);
    for (const [code] of segments) be(sub, code, 2);
    be(sub, 0, 2);
    for (const [code] of segments) be(sub, code, 2);
    for (const [code, gid] of segments) be(sub, code === 0xffff ? 1 : (gid - code) & 0xffff, 2);
    for (let i = 0; i < segments.length; i += 1) be(sub, 0, 2);
    const table: number[] = [];
    be(table, 0, 2);
    be(table, 1, 2);
    be(table, 3, 2);
    be(table, 1, 2);
    be(table, 12, 4);
    tables.unshift(["cmap", [...table, ...sub]]);
  }
  const directory: number[] = [];
  be(directory, 0x00010000, 4);
  be(directory, tables.length, 2);
  for (let i = 0; i < 3; i += 1) be(directory, 0, 2);
  let at = 12 + tables.length * 16;
  const body: number[] = [];
  for (const [tag, bytes] of tables) {
    for (const ch of tag) directory.push(ch.charCodeAt(0));
    be(directory, 0, 4);
    be(directory, at, 4);
    be(directory, bytes.length, 4);
    body.push(...bytes);
    while (body.length % 4) body.push(0);
    at = 12 + tables.length * 16 + body.length;
  }
  return new Uint8Array([...directory, ...body]);
}

const BAR: Outline = { kind: "simple", contours: [[[0, 0], [100, 0], [100, 700], [0, 700]]] };
const WEDGE: Outline = { kind: "simple", contours: [[[0, 0], [500, 0], [250, 700]], [[200, 100], [300, 100], [250, 300]]] };
const HOOK: Outline = { kind: "simple", contours: [[[0, 0], [400, 0], [400, 300], [100, 300], [100, 700]]] };
const ODD: Outline = { kind: "simple", contours: [[[0, 0], [10, 990], [20, 0]]] };
const bars = (gid: number): Outline => ({ kind: "composite", parts: [{ gid, dx: 0, dy: 0 }, { gid, dx: 300, dy: 0 }] });

/** The reference font: I is a bar, A a wedge, J a hook, H two bars, a space has no outline. */
const REFERENCE = makeFont([{ kind: "empty" }, BAR, WEDGE, HOOK, bars(1), { kind: "empty" }], [[0x41, 2], [0x48, 4], [0x49, 1], [0x4a, 3], [0x20, 5]]);

/** An embedded subset: the same outlines in first-use order, reachable only through PDF.js's PUA codes. */
const EMBEDDED = makeFont([{ kind: "empty" }, HOOK, BAR, WEDGE, bars(2), ODD], [[0xe001, 1], [0xe002, 2], [0xe003, 3], [0xe004, 4], [0xe005, 5]]);
const TO_FONT_CHAR = new Map([[1, 0xe001], [2, 0xe002], [3, 0xe003], [4, 0xe004], [5, 0xe005]]);

describe("glyph outlines name the letters of a font with no letter map", () => {
  it("reads the format 4 character map and fingerprints simple and composite glyphs", () => {
    const font = readGlyphFont(REFERENCE)!;
    expect(font.codes()).toEqual([[0x20, 5], [0x41, 2], [0x48, 4], [0x49, 1], [0x4a, 3]]);
    expect(font.glyphFor(0x41)).toBe(2);
    expect(font.glyphFor(0x42)).toBeNull();
    expect(font.fingerprint(1)).toMatch(/^[0-9a-f]{16}$/);
    expect(font.fingerprint(4)).not.toBe(font.fingerprint(1));
    expect(font.fingerprint(9)).toBeNull();
    expect(readGlyphFont(new Uint8Array([1, 2, 3]))).toBeNull();
  });

  it("matches the embedded glyphs by outline, with a space for an empty glyph", () => {
    const { table } = fingerprintTable([readGlyphFont(REFERENCE)!], (code) => code > 0x20);
    const letters = glyphLetters(EMBEDDED, TO_FONT_CHAR, table)!;
    expect([1, 2, 3, 4].map(letters).join("")).toBe("JIAH");
    expect(letters(5)).toBe(UNKNOWN_GLYPH);
    expect(letters(9)).toBe(" ");
  });

  it("packs and unpacks the table, and the Arial table covers letters and digits", () => {
    const { table } = fingerprintTable([readGlyphFont(REFERENCE)!], () => true);
    expect(unpackFingerprints(packFingerprints(table))).toEqual(table);
    expect(ARIAL_GLYPHS.length % 17).toBe(0);
    const letters = new Set(unpackFingerprints(ARIAL_GLYPHS).values());
    for (const ch of "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789.,-/()") expect(letters.has(ch)).toBe(true);
  });

  it("calls text unreadable only when glyph numbers or unknown glyphs pass 2%", () => {
    expect(unreadablePdfText("Plan Sections 600.00 90.00")).toBe(false);
    expect(unreadablePdfText("\u0001\u0002\u0003 \u0004")).toBe(true);
    expect(unreadablePdfText(`${"A".repeat(99)}${UNKNOWN_GLYPH}`)).toBe(false);
    expect(unreadablePdfText(`${"A".repeat(90)}${UNKNOWN_GLYPH.repeat(10)}`)).toBe(true);
    expect(unreadablePdfText("   ")).toBe(false);
  });
});

// ---------------------------------------------------------------- glyphs PDF.js left out of the text

const OPS: PdfTextOps = { save: 10, restore: 11, setFont: 37, showText: 44, showSpacedText: 45, nextLineShowText: 46, nextLineSetSpacingShowText: 47 };
const glyph = (code: number) => ({ originalCharCode: code, unicode: "?" });
const run = (codes: number[], at: number): DrawnGlyph[] => codes.map((code) => ({ code, run: at }));
const item = (str: string, x: number, width: number, y = 700): PdfTextItem => ({ str, transform: [10, 0, 0, 10, x, y], width, height: 10, fontName: "F1" });
const codes = (...list: number[]) => String.fromCharCode(...list);

describe("left-out glyphs go back where they were drawn", () => {
  it("follows the font through save and restore and reads every show-text form", () => {
    const drawn = drawnGlyphs({
      fnArray: [37, 44, 10, 37, 45, 11, 46, 47],
      argsArray: [["F1", 10], [[glyph(1), glyph(2)]], null, ["F2", 10], [[glyph(3), 120, glyph(4)]], null, [[glyph(5)]], [0, 0, [glyph(6)]]],
    }, OPS);
    expect(drawn.get("F1")).toEqual([{ code: 1, run: 1 }, { code: 2, run: 1 }, { code: 5, run: 6 }, { code: 6, run: 7 }]);
    expect(drawn.get("F2")).toEqual([{ code: 3, run: 4 }, { code: 4, run: 4 }]);
  });

  // 9 to 13 and 32 read as whitespace to PDF.js, so it drops them; here they are e, the space, r and m.
  const letters: Letters = (code) => ({ 1: "c", 2: "n", 3: "t", 4: "o", 7: " ", 9: "e", 10: " ", 13: "r", 20: "H", 21: "l", 32: "m", 33: "s" })[code] ?? "?";
  const fonts = new Map([["F1", letters]]);

  it("puts the letters in place of the space PDF.js left, or before the next character", () => {
    const drawn = new Map([["F1", run([1, 9, 2, 3, 9, 13, 10, 3, 4], 1)]]);
    const out = alignGlyphs([item(codes(1, 32, 2, 3, 32, 3, 4), 0, 40)], drawn, fonts);
    expect(out[0]!.str).toBe("center to");
    const inline = alignGlyphs([item(codes(1, 2, 3), 0, 20)], new Map([["F1", run([1, 9, 2, 3], 1)]]), fonts);
    expect(inline[0]!.str).toBe("cent");
  });

  it("joins a blank item that ends a word to the item before it", () => {
    const drawn = new Map([["F1", run([1, 9, 2, 3, 9, 13, 10, 3, 4], 1)]]);
    const out = alignGlyphs([item(codes(1, 32, 2, 3), 0, 20), item(" ", 20, 1), item(codes(3, 4), 30, 10)], drawn, fonts);
    expect(out.map((i) => i.str)).toEqual(["center ", "", "to"]);
  });

  it("joins a blank item that starts a word to the item after it", () => {
    const drawn = new Map([["F1", run([2, 7, 32, 33, 33], 1)]]);
    const out = alignGlyphs([item(codes(2, 7), 0, 10), item(" ", 10, 1), item(codes(33, 33), 18, 10)], drawn, fonts);
    expect(out.map((i) => i.str)).toEqual(["n ", "", "mss"]);
  });

  it("makes one item of a word split around a blank item", () => {
    const drawn = new Map([["F1", run([20, 9, 9, 21], 1)]]);
    const out = alignGlyphs([item(codes(20), 0, 8), item(" ", 8, 1), item(codes(21), 20, 3)], drawn, fonts);
    expect(out.map((i) => i.str)).toEqual(["Heel", "", ""]);
    expect(out[0]!.width).toBe(23);
  });

  it("leaves items in other fonts as they are", () => {
    const plain = { ...item("Plan Sections", 0, 60), fontName: "F9" };
    expect(alignGlyphs([plain], new Map([["F1", run([1], 1)]]), fonts)[0]!.str).toBe("Plan Sections");
  });
});

// ---------------------------------------------------------------- whole documents

function fakePage(items: PdfTextItem[], ops: { fnArray: number[]; argsArray: unknown[] }, objects: Record<string, unknown>, calls: string[]): PdfPageLike {
  return {
    getTextContent: async (params) => {
      calls.push(params?.disableNormalization ? "text raw" : "text");
      return { items };
    },
    getOperatorList: async () => {
      calls.push("ops");
      return ops;
    },
    commonObjs: { has: (id) => id in objects, get: (id) => objects[id] },
    cleanup: () => undefined,
  };
}

describe("pdfDocumentText", () => {
  it("decodes a page whose font gives glyph numbers and reads other pages as before", async () => {
    const { table } = fingerprintTable([readGlyphFont(REFERENCE)!], (code) => code > 0x20);
    const calls: string[] = [];
    const pages = [
      fakePage([item(codes(3, 2), 10, 20)], { fnArray: [37, 44], argsArray: [["F1", 10], [[glyph(3), glyph(2), glyph(10)]]] }, { F1: { data: EMBEDDED, toFontChar: new Map([...TO_FONT_CHAR, [10, 0xe002]]) } }, calls),
      fakePage([{ ...item("Summary", 10, 40), fontName: "F2" }], { fnArray: [], argsArray: [] }, {}, calls),
    ];
    const doc: PdfDocLike = { numPages: 2, getPage: async (n) => pages[n - 1]! };
    const text = await pdfDocumentText(doc, OPS, table);
    expect(text.split("\n\f").map((page) => page.trim())).toEqual(["AII", "Summary"]);
    expect(calls).toEqual(["text", "text raw", "ops", "text"]);
  });
});

// ---------------------------------------------------------------- what the imports say

function makeBook(): SheetBook {
  return new SheetBook({ sheets: { s0: { id: "s0", name: "Setup", cellData: {} }, s1: { id: "s1", name: "Data", cellData: {} } }, styles: {} } as never, []);
}

describe("imports say plainly when a PDF cannot be read", () => {
  const glyphNumbers = "\u0001\u0002\u0003\u0004 \u0005\u0006\u0007\u0008\n".repeat(20);

  it("Import Plan names the glyph-number fonts instead of a missing table", async () => {
    const book = makeBook();
    const result = await importSurveyPlan(book, { path: "plan.pdf", kind: "pdf", text: glyphNumbers }, null);
    expect(result.ok).toBe(false);
    expect(result.alert).toContain("glyph numbers");
    expect(book.writes).toEqual([]);
    const noTable = await importSurveyPlan(book, { path: "plan.pdf", kind: "pdf", text: "Planning Report\nnothing here" }, null);
    expect(noTable.alert).not.toContain("glyph numbers");
  });

  it("Import AC says so for glyph numbers, and reports a readable report with no wells", () => {
    const unreadable = importAntiCollision(makeBook(), "ac.pdf", glyphNumbers, null);
    expect(unreadable.ok).toBe(false);
    expect(unreadable.alert).toContain("glyph numbers");
    const empty = importAntiCollision(makeBook(), "ac.pdf", "Anticollision Report\nSummary\nno rows\n", null);
    expect(empty).toMatchObject({ ok: true, status: "AC import complete: 0 offset well(s), 0 with SF < 2.0." });
    expect(empty.note).toContain("gave no offset wells");
    expect(importAntiCollision(makeBook(), "ac.pdf", "Summary\nno rows\n", null, 1e6, true).note).toBeUndefined();
  });

  it("reads a Summary wellbore whose event number has one digit", async () => {
    const result = importAntiCollision(makeBook(), "ac.pdf", "Summary\n  100/01-02-003-04W5/0 - Offset D   1,200.00   1,190.00   20.00   12.00   1.800 SF\n", null);
    expect(result.hits.map((hit) => hit.well)).toEqual(["100/01-02-003-04W5/0 - Offset D"]);
  });
});
