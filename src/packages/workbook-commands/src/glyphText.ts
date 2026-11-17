import { ARIAL_GLYPHS } from "./arialGlyphs";
import { glyphLetters, UNKNOWN_GLYPH, unpackFingerprints } from "./glyphShape";
import { layoutText, type PdfTextItem } from "./pdfLayout";

// PDF text for fonts that give glyph numbers instead of letters (glyphShape.ts). PDF.js's text
// content drops every glyph whose number reads as whitespace, and in such a font 9 to 13 and 32
// are letters. So the glyph numbers come from the page's operator list, which keeps every glyph in
// drawing order, and are laid over the text items, which keep the positions.

/** The font fields PDF.js keeps with `fontExtraProperties`: the rebuilt font and its code map. */
interface PdfFontObject {
  data?: Uint8Array;
  toFontChar?: unknown;
}

export interface PdfOperatorList {
  fnArray: number[];
  argsArray: unknown[];
}

/** The PDF.js `OPS` numbers that save or restore state, set the font and show text. */
export interface PdfTextOps {
  save: number;
  restore: number;
  setFont: number;
  showText: number;
  showSpacedText: number;
  nextLineShowText: number;
  nextLineSetSpacingShowText: number;
}

/** The parts of a PDF.js page and document that `pdfDocumentText` reads. */
export interface PdfPageLike {
  getTextContent(params?: { disableNormalization?: boolean }): Promise<{ items: unknown[] }>;
  getOperatorList(): Promise<PdfOperatorList>;
  commonObjs: { has(id: string): boolean; get(id: string): unknown };
  cleanup(): unknown;
}

export interface PdfDocLike {
  numPages: number;
  getPage(page: number): Promise<PdfPageLike>;
}

export type Letters = (code: number) => string;

/** One drawn glyph: its number in the font, and the show-text operation that drew it. */
export interface DrawnGlyph {
  code: number;
  run: number;
}

/** Control codes no font's real text contains: a font with them is giving glyph numbers. */
const GLYPH_NUMBER = /[\u0000-\u0008\u000b\u000e-\u001f]/;
/** What PDF.js leaves out of text content: whitespace and invisible format marks. */
const LEFT_OUT = /[\s\u00ad\u061c\u180e\u200b-\u200f\u2060-\u2064\ufeff]/u;
const SEARCH_AHEAD = 32;

let arialTable: Map<string, string> | null = null;

function arialGlyphTable(): Map<string, string> {
  arialTable ??= unpackFingerprints(ARIAL_GLYPHS);
  return arialTable;
}

/** True when more than 2% of the printed characters are glyph numbers or unmatched glyphs. */
export function unreadablePdfText(text: string): boolean {
  let shown = 0;
  let bad = 0;
  for (const ch of text) {
    if (/\s/.test(ch)) continue;
    shown += 1;
    if (ch === UNKNOWN_GLYPH || GLYPH_NUMBER.test(ch)) bad += 1;
  }
  return shown > 0 && bad / shown > 0.02;
}

function leftOut(code: number): boolean {
  return LEFT_OUT.test(String.fromCodePoint(code));
}

/** Each font's glyphs in drawing order. The font is part of the state that save and restore keep. */
export function drawnGlyphs(list: PdfOperatorList, ops: PdfTextOps): Map<string, DrawnGlyph[]> {
  const out = new Map<string, DrawnGlyph[]>();
  const saved: Array<string | null> = [];
  let font: string | null = null;
  list.fnArray.forEach((fn, i) => {
    const args = list.argsArray[i];
    if (fn === ops.save) saved.push(font);
    else if (fn === ops.restore) font = saved.length > 0 ? saved.pop()! : font;
    else if (fn === ops.setFont) font = Array.isArray(args) && typeof args[0] === "string" ? args[0] : null;
    else if (fn === ops.showText || fn === ops.showSpacedText || fn === ops.nextLineShowText || fn === ops.nextLineSetSpacingShowText) {
      const glyphs = Array.isArray(args) ? args[fn === ops.nextLineSetSpacingShowText ? 2 : 0] : null;
      if (font == null || !Array.isArray(glyphs)) return;
      const drawn = out.get(font) ?? [];
      out.set(font, drawn);
      for (const glyph of glyphs) {
        const code = glyph != null && typeof glyph === "object" ? (glyph as { originalCharCode?: unknown }).originalCharCode : null;
        if (typeof code === "number") drawn.push({ code, run: i });
      }
    }
  });
  return out;
}

/**
 * Text items with the letters of every glyph drawn for them. Each visible character is matched to
 * its glyph in drawing order. A glyph PDF.js left out goes back where it was drawn: in place of the
 * space PDF.js put there, or before the next character, or at the end of the item before when it
 * was drawn with that item's last character.
 */
export function alignGlyphs(items: PdfTextItem[], drawn: ReadonlyMap<string, DrawnGlyph[]>, letters: ReadonlyMap<string, Letters | null>): PdfTextItem[] {
  const parts = items.map(() => [] as string[]);
  const spans = items.map(() => null as { first: number; last: number } | null);
  const cursor = new Map<string, number>();
  const last = new Map<string, { item: number; run: number }>();
  const took = (index: number, first: number, end: number) => {
    const span = spans[index];
    spans[index] = span ? { first: span.first, last: end } : { first, last: end };
  };
  items.forEach((item, index) => {
    const font = item.fontName;
    const decode = font ? letters.get(font) : null;
    const glyphs = font ? drawn.get(font) : undefined;
    const out = parts[index]!;
    if (!font || !decode || !glyphs) {
      out.push(item.str);
      return;
    }
    let at = cursor.get(font) ?? 0;
    let prev = last.get(font) ?? null;
    for (const ch of item.str) {
      if (ch === " ") {
        const start = at;
        while (at < glyphs.length && leftOut(glyphs[at]!.code)) at += 1;
        if (at === start) {
          out.push(" ");
          continue;
        }
        for (let k = start; k < at; k += 1) out.push(decode(glyphs[k]!.code));
        took(index, start, at - 1);
        prev = { item: index, run: glyphs[at - 1]!.run };
        continue;
      }
      const code = ch.codePointAt(0)!;
      let match = at;
      while (match < glyphs.length && match - at < SEARCH_AHEAD && glyphs[match]!.code !== code) match += 1;
      if (match >= glyphs.length || glyphs[match]!.code !== code) {
        out.push(decode(code));
        continue;
      }
      const run = glyphs[match]!.run;
      for (let k = at; k < match; k += 1) {
        const glyph = glyphs[k]!;
        if (!leftOut(glyph.code)) continue;
        if (prev && prev.item !== index && glyph.run === prev.run && glyph.run !== run) parts[prev.item]!.push(decode(glyph.code));
        else out.push(decode(glyph.code));
      }
      out.push(decode(code));
      took(index, at, match);
      prev = { item: index, run };
      at = match + 1;
    }
    cursor.set(font, at);
    if (prev) last.set(font, prev);
  });
  for (const [font, glyphs] of drawn) {
    const end = last.get(font);
    const decode = letters.get(font);
    if (!end || !decode) continue;
    for (let k = cursor.get(font) ?? 0; k < glyphs.length; k += 1) {
      const glyph = glyphs[k]!;
      if (glyph.run === end.run && leftOut(glyph.code)) parts[end.item]!.push(decode(glyph.code));
    }
  }
  const joined = joinLoneGlyphs(items, parts, spans);
  return joined.map((item, index) => ({ ...item, str: parts[index]!.join("") }));
}

/**
 * PDF.js can give left-out glyphs a blank item of their own, which the layout would set apart.
 * Their letters join the item drawn just before them when they end a word, the item drawn just
 * after when they start one, and inside a word the two items on either side become one.
 */
function joinLoneGlyphs(items: PdfTextItem[], parts: string[][], spans: Array<{ first: number; last: number } | null>): PdfTextItem[] {
  const out = items.map((item) => ({ ...item }));
  const near = (index: number, step: 1 | -1, glyph: number): number => {
    const font = items[index]!.fontName;
    for (let j = index + step; j >= 0 && j < items.length; j += step) {
      const span = spans[j];
      if (items[j]!.fontName !== font || !span) continue;
      return (step === 1 ? span.first : span.last) === glyph ? j : -1;
    }
    return -1;
  };
  items.forEach((item, index) => {
    const span = spans[index];
    const text = parts[index]!.join("");
    if (!span || item.str.trim() !== "" || text.trim() === "") return;
    const next = near(index, 1, span.last + 1);
    const before = near(index, -1, span.first - 1);
    if (next < 0 && before < 0) return;
    const nextText = next >= 0 ? parts[next]!.join("") : "";
    const sameLine = before >= 0 && next >= 0 && Math.abs(out[before]!.transform[5]! - out[next]!.transform[5]!) < 0.5;
    if (before >= 0 && (next < 0 || /\s$/.test(text))) {
      parts[before]!.push(text);
      spans[before] = { first: spans[before]!.first, last: span.last };
    } else if (next >= 0 && (before < 0 || /^\s/.test(text) || /\s$/.test(parts[before]!.join("")) || !sameLine)) {
      parts[next] = [text, nextText];
      spans[next] = { first: span.first, last: spans[next]!.last };
    } else {
      parts[before]!.push(text, nextText);
      spans[before] = { first: spans[before]!.first, last: spans[next]!.last };
      out[before]!.width = out[next]!.transform[4]! + out[next]!.width - out[before]!.transform[4]!;
      parts[next] = [];
      spans[next] = null;
    }
    parts[index] = [];
    spans[index] = null;
  });
  return out;
}

function textItem(item: unknown): PdfTextItem[] {
  if (item == null || typeof item !== "object" || !("str" in item)) return [];
  const raw = item as { str: string; transform: number[]; width: number; height: number; fontName?: string };
  return [{ str: raw.str, transform: raw.transform, width: raw.width, height: raw.height, ...(raw.fontName ? { fontName: raw.fontName } : {}) }];
}

function fontLetters(page: PdfPageLike, id: string, table: ReadonlyMap<string, string>): Letters | null {
  if (!page.commonObjs.has(id)) return null;
  const font = page.commonObjs.get(id) as PdfFontObject | null;
  if (!font?.data || !(font.toFontChar instanceof Map)) return null;
  return glyphLetters(font.data, font.toFontChar as ReadonlyMap<number, number>, table);
}

/**
 * The document's text in layout order. A page whose text has glyph numbers is read again without
 * Unicode normalisation, and its operator list is fetched for the glyphs and the fonts. Other pages
 * read as before.
 */
export async function pdfDocumentText(doc: PdfDocLike, ops: PdfTextOps, table: ReadonlyMap<string, string> = arialGlyphTable()): Promise<string> {
  const pages: PdfTextItem[][] = [];
  const letters = new Map<string, Letters | null>();
  for (let p = 1; p <= doc.numPages; p += 1) {
    const page = await doc.getPage(p);
    let items = (await page.getTextContent()).items.flatMap(textItem);
    const fonts = new Set<string>();
    for (const item of items) {
      if (!item.fontName) continue;
      if (letters.get(item.fontName) || (!letters.has(item.fontName) && GLYPH_NUMBER.test(item.str))) fonts.add(item.fontName);
    }
    if (fonts.size > 0) {
      items = (await page.getTextContent({ disableNormalization: true })).items.flatMap(textItem);
      const drawn = drawnGlyphs(await page.getOperatorList(), ops);
      for (const id of fonts) if (!letters.has(id)) letters.set(id, fontLetters(page, id, table));
      items = alignGlyphs(items, drawn, letters);
    }
    pages.push(items);
    page.cleanup();
  }
  return layoutText(pages);
}
