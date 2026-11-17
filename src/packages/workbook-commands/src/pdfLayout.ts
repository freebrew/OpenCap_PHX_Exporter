// Text for the Import Plan / Import AC parsers, rebuilt from PDF.js text items so it reads
// like `pdftotext -layout` (the extractor the VBA prefers): one line per baseline, items
// left to right, and a column gap kept as spaces. Rotated text (plot labels) is left out.

export interface PdfTextItem {
  str: string;
  /** [a, b, c, d, e, f]: e/f are the baseline origin in PDF points, y up. */
  transform: number[];
  width: number;
  height: number;
  /** PDF.js's id for the item's font. */
  fontName?: string;
}

interface Placed {
  text: string;
  x: number;
  y: number;
  right: number;
  size: number;
}

function placed(item: PdfTextItem): Placed | null {
  const [a = 0, b = 0, c = 0, d = 0, e = 0, f = 0] = item.transform;
  if (item.str.trim() === "") return null;
  if (Math.abs(b) > 1e-3 || Math.abs(c) > 1e-3 || a <= 0) return null;
  const size = Math.abs(d) || item.height || a;
  return { text: item.str, x: e, y: f, right: e + item.width, size };
}

function median(values: number[]): number {
  if (values.length === 0) return 0;
  const sorted = [...values].sort((p, q) => p - q);
  return sorted[Math.floor(sorted.length / 2)]!;
}

function pageText(items: PdfTextItem[]): string {
  const glyphs = items.flatMap((item) => {
    const one = placed(item);
    return one ? [one] : [];
  });
  if (glyphs.length === 0) return "";
  const pitch = median(glyphs.map((g) => (g.right - g.x) / Math.max(1, g.text.length)).filter((w) => w > 0)) || 4;
  glyphs.sort((p, q) => q.y - p.y || p.x - q.x);

  const lines: Placed[][] = [];
  for (const glyph of glyphs) {
    const line = lines[lines.length - 1];
    const base = line?.[0];
    if (line && base && Math.abs(base.y - glyph.y) <= 0.45 * Math.max(base.size, glyph.size)) line.push(glyph);
    else lines.push([glyph]);
  }

  return lines.map((line) => {
    line.sort((p, q) => p.x - q.x);
    let out = "";
    let end = Number.NEGATIVE_INFINITY;
    for (const glyph of line) {
      const gap = glyph.x - end;
      if (out !== "" && gap > 0.15 * glyph.size) out += " ".repeat(Math.max(1, Math.round(gap / pitch)));
      out += glyph.text;
      end = Math.max(end, glyph.right);
    }
    return out.replace(/\s+$/, "");
  }).join("\n");
}

/** Pages joined with a form feed, as pdftotext separates them. */
export function layoutText(pages: PdfTextItem[][]): string {
  return pages.map(pageText).join("\n\f");
}
