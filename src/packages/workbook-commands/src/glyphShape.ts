// Letters for PDF text whose font has no letter map. Some report exports (Crystal Reports through
// Acrobat Distiller) embed Arial as a composite font with neither a ToUnicode table nor a cmap, so
// every reader, PDFium and PDF.js alike, gets glyph numbers instead of text. The glyph outlines are
// still there, and they are Arial's: a glyph is named by matching its outline fingerprint to a
// table taken from the Arial fonts. No outlines are stored, only fingerprints.

/** A TrueType font, read far enough to fingerprint its glyphs and follow its character map. */
export interface GlyphFont {
  /** The outline fingerprint of a glyph, or null when the glyph cannot be read. */
  fingerprint(gid: number): string | null;
  /** The glyph the font's own character map gives for a character code, or null. */
  glyphFor(code: number): number | null;
  /** Every character code the font's character map covers, with its glyph, in code order. */
  codes(): Array<[number, number]>;
}

export const UNKNOWN_GLYPH = "\uFFFD";
const FINGERPRINT_CHARS = 16;
const MAX_COMPONENT_DEPTH = 8;

interface Table {
  offset: number;
  length: number;
}

function tableDirectory(view: DataView): Map<string, Table> {
  const tables = new Map<string, Table>();
  const count = view.getUint16(4);
  for (let i = 0; i < count; i += 1) {
    const at = 12 + i * 16;
    const tag = String.fromCharCode(view.getUint8(at), view.getUint8(at + 1), view.getUint8(at + 2), view.getUint8(at + 3));
    const offset = view.getUint32(at + 8);
    const length = view.getUint32(at + 12);
    if (offset + length <= view.byteLength) tables.set(tag, { offset, length });
  }
  return tables;
}

/** FNV-1a run twice with different seeds, as 16 hex characters. */
function hashText(text: string): string {
  let a = 0x811c9dc5;
  let b = 0x9747b28c;
  for (let i = 0; i < text.length; i += 1) {
    const c = text.charCodeAt(i);
    a = Math.imul(a ^ c, 0x01000193);
    b = Math.imul(b ^ c, 0x5bd1e995);
    b ^= b >>> 15;
  }
  return (a >>> 0).toString(16).padStart(8, "0") + (b >>> 0).toString(16).padStart(8, "0");
}

function cmapEntries(view: DataView, cmap: Table): Map<number, number> {
  const map = new Map<number, number>();
  const count = view.getUint16(cmap.offset + 2);
  for (let i = 0; i < count; i += 1) {
    const sub = cmap.offset + view.getUint32(cmap.offset + 4 + i * 8 + 4);
    if (sub + 4 > cmap.offset + cmap.length) continue;
    const format = view.getUint16(sub);
    if (format === 4) readFormat4(view, sub, map);
    else if (format === 12) readFormat12(view, sub, map);
  }
  return map;
}

function readFormat4(view: DataView, sub: number, map: Map<number, number>): void {
  const segX2 = view.getUint16(sub + 6);
  const ends = sub + 14;
  const starts = ends + segX2 + 2;
  const deltas = starts + segX2;
  const ranges = deltas + segX2;
  for (let s = 0; s < segX2; s += 2) {
    const end = view.getUint16(ends + s);
    const start = view.getUint16(starts + s);
    const delta = view.getInt16(deltas + s);
    const rangeOffset = view.getUint16(ranges + s);
    for (let code = start; code <= end && code !== 0xffff; code += 1) {
      let gid: number;
      if (rangeOffset === 0) gid = (code + delta) & 0xffff;
      else {
        const at = ranges + s + rangeOffset + 2 * (code - start);
        if (at + 2 > view.byteLength) continue;
        const raw = view.getUint16(at);
        gid = raw === 0 ? 0 : (raw + delta) & 0xffff;
      }
      if (gid !== 0 && !map.has(code)) map.set(code, gid);
    }
  }
}

function readFormat12(view: DataView, sub: number, map: Map<number, number>): void {
  const groups = view.getUint32(sub + 12);
  for (let g = 0; g < groups; g += 1) {
    const at = sub + 16 + g * 12;
    const start = view.getUint32(at);
    const end = Math.min(view.getUint32(at + 4), start + 0xffff);
    const first = view.getUint32(at + 8);
    for (let code = start; code <= end; code += 1) if (!map.has(code)) map.set(code, first + code - start);
  }
}

/** Reads the head, loca, glyf and cmap tables. Null when the bytes are not a TrueType font with outlines. */
export function readGlyphFont(bytes: Uint8Array): GlyphFont | null {
  if (bytes.byteLength < 12) return null;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const tag = view.getUint32(0);
  if (tag !== 0x00010000 && tag !== 0x74727565) return null;
  const tables = tableDirectory(view);
  const head = tables.get("head");
  const maxp = tables.get("maxp");
  const loca = tables.get("loca");
  const glyf = tables.get("glyf");
  if (!head || !maxp || !loca || !glyf) return null;
  const longOffsets = view.getInt16(head.offset + 50) === 1;
  const numGlyphs = view.getUint16(maxp.offset + 4);
  const glyphStart = (gid: number) => glyf.offset + (longOffsets ? view.getUint32(loca.offset + gid * 4) : view.getUint16(loca.offset + gid * 2) * 2);
  const cmapTable = tables.get("cmap");
  let map: Map<number, number> | null = null;
  const cmap = () => (map ??= cmapTable ? cmapEntries(view, cmapTable) : new Map());
  const outlines = new Map<number, string | null>();

  const simpleOutline = (at: number, contours: number): string => {
    const ends: number[] = [];
    for (let i = 0; i < contours; i += 1) ends.push(view.getUint16(at + 10 + i * 2));
    let p = at + 10 + contours * 2;
    p += 2 + view.getUint16(p);
    const count = contours > 0 ? ends[contours - 1]! + 1 : 0;
    const flags = new Uint8Array(count);
    for (let i = 0; i < count;) {
      const flag = view.getUint8(p);
      p += 1;
      flags[i] = flag;
      i += 1;
      if (flag & 8) {
        let repeat = view.getUint8(p);
        p += 1;
        while (repeat > 0 && i < count) {
          flags[i] = flag;
          i += 1;
          repeat -= 1;
        }
      }
    }
    const coords = (shortBit: number, sameBit: number): number[] => {
      const out: number[] = [];
      let value = 0;
      for (let i = 0; i < count; i += 1) {
        const flag = flags[i]!;
        if (flag & shortBit) {
          const step = view.getUint8(p);
          p += 1;
          value += flag & sameBit ? step : -step;
        } else if (!(flag & sameBit)) {
          value += view.getInt16(p);
          p += 2;
        }
        out.push(value);
      }
      return out;
    };
    const xs = coords(2, 16);
    const ys = coords(4, 32);
    const points = xs.map((x, i) => `${x},${ys[i]},${flags[i]! & 1}`);
    return `${ends.join(",")}|${points.join(";")}`;
  };

  const compositeOutline = (at: number, depth: number): string | null => {
    const parts: string[] = [];
    let p = at + 10;
    for (;;) {
      const flags = view.getUint16(p);
      const component = view.getUint16(p + 2);
      p += 4;
      let dx: number;
      let dy: number;
      if (flags & 1) {
        dx = view.getInt16(p);
        dy = view.getInt16(p + 2);
        p += 4;
      } else {
        dx = view.getInt8(p);
        dy = view.getInt8(p + 1);
        p += 2;
      }
      const scale: number[] = [];
      const words = flags & 8 ? 1 : flags & 0x40 ? 2 : flags & 0x80 ? 4 : 0;
      for (let w = 0; w < words; w += 1) scale.push(view.getInt16(p + w * 2));
      p += words * 2;
      const inner = outline(component, depth + 1);
      if (inner == null) return null;
      parts.push(`${inner}@${dx},${dy},${flags & 2 ? "xy" : "pt"}${scale.length ? `:${scale.join(",")}` : ""}`);
      if (!(flags & 0x20)) break;
    }
    return `c[${parts.join("+")}]`;
  };

  const outline = (gid: number, depth: number): string | null => {
    if (gid < 0 || gid >= numGlyphs || depth > MAX_COMPONENT_DEPTH) return null;
    const known = outlines.get(gid);
    if (known !== undefined) return known;
    let result: string | null = null;
    try {
      const start = glyphStart(gid);
      const end = glyphStart(gid + 1);
      if (end <= start) result = "";
      else {
        const contours = view.getInt16(start);
        result = contours >= 0 ? simpleOutline(start, contours) : compositeOutline(start, depth);
      }
    } catch {
      result = null;
    }
    outlines.set(gid, result);
    return result;
  };

  return {
    fingerprint: (gid) => {
      const text = outline(gid, 0);
      return text == null ? null : hashText(text);
    },
    glyphFor: (code) => cmap().get(code) ?? null,
    codes: () => [...cmap().entries()].sort((a, b) => a[0] - b[0]),
  };
}

/** The packed table: each entry is a fingerprint followed by its character. */
export function packFingerprints(entries: Map<string, string>): string {
  return [...entries.entries()].sort((a, b) => (a[0] < b[0] ? -1 : 1)).map(([key, ch]) => key + ch).join("");
}

export function unpackFingerprints(packed: string): Map<string, string> {
  const map = new Map<string, string>();
  const step = FINGERPRINT_CHARS + 1;
  for (let i = 0; i + step <= packed.length; i += step) map.set(packed.slice(i, i + FINGERPRINT_CHARS), packed[i + FINGERPRINT_CHARS]!);
  return map;
}

/**
 * Fingerprint to character for the given reference fonts. The lowest code wins when two characters
 * share an outline (a space and a no-break space), so the plain ASCII form comes out.
 */
export function fingerprintTable(fonts: GlyphFont[], keep: (code: number) => boolean): { table: Map<string, string>; clashes: string[] } {
  const table = new Map<string, string>();
  const clashes: string[] = [];
  for (const font of fonts) {
    for (const [code, gid] of font.codes()) {
      if (!keep(code)) continue;
      const key = font.fingerprint(gid);
      if (key == null) continue;
      const ch = String.fromCharCode(code);
      const known = table.get(key);
      if (known == null) table.set(key, ch);
      else if (known !== ch) clashes.push(`${known}=${ch}`);
    }
  }
  return { table, clashes };
}

/**
 * The letters of an embedded font whose text came out as glyph numbers. `toFontChar` is PDF.js's
 * map from a text code to the character its rebuilt font draws that glyph with. A code PDF.js left
 * out of that map has no outline, which is a space. Null when the font cannot be read.
 */
export function glyphLetters(fontData: Uint8Array, toFontChar: ReadonlyMap<number, number>, table: ReadonlyMap<string, string>): ((code: number) => string) | null {
  const font = readGlyphFont(fontData);
  if (!font) return null;
  const letters = new Map<number, string>();
  return (code) => {
    const known = letters.get(code);
    if (known != null) return known;
    const fontChar = toFontChar.get(code);
    let ch = " ";
    if (fontChar != null) {
      const gid = font.glyphFor(fontChar);
      const key = gid == null ? null : font.fingerprint(gid);
      ch = key == null ? UNKNOWN_GLYPH : (table.get(key) ?? UNKNOWN_GLYPH);
    }
    letters.set(code, ch);
    return ch;
  };
}
