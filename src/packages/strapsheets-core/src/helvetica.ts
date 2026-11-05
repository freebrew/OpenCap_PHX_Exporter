// Widths of the PDF standard fonts Helvetica and Helvetica-Bold (Adobe AFM, 1000 units to the
// em), so the card can centre, clip and wrap text without embedding a font. ASCII only: a
// character outside it is drawn as "?" (see pdfString), and measured as one.

// Characters 32..126, in order.
const REGULAR = [
  278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278,
  556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 278, 278, 584, 584, 584, 556,
  1015, 667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556, 833, 722, 778,
  667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 278, 278, 278, 469, 556,
  333, 556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500, 222, 833, 556, 556,
  556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584,
];

const BOLD = [
  278, 333, 474, 556, 556, 889, 722, 238, 333, 333, 389, 584, 278, 333, 278, 278,
  556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 333, 333, 584, 584, 584, 611,
  975, 722, 722, 722, 722, 667, 611, 778, 722, 278, 556, 722, 611, 833, 722, 778,
  667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 333, 278, 333, 584, 556,
  333, 556, 611, 556, 611, 556, 333, 611, 611, 278, 278, 556, 278, 889, 611, 611,
  611, 611, 389, 556, 333, 611, 556, 778, 556, 556, 500, 389, 280, 389, 584,
];

/** The text as the PDF draws it: printable ASCII, anything else a question mark. */
export function plainText(text: string): string {
  return text.replace(/[^\x20-\x7e]/g, "?");
}

function advance(character: string, bold: boolean): number {
  const code = character.charCodeAt(0);
  return (bold ? BOLD : REGULAR)[code >= 32 && code <= 126 ? code - 32 : 31] ?? 556;
}

/** Width of the text at a font size, in the same unit as the size. */
export function textWidth(text: string, bold: boolean, size: number): number {
  let units = 0;
  for (const character of plainText(text)) units += advance(character, bold);
  return (units * size) / 1000;
}

/**
 * The units a line may break between: words, except that a bracketed thread size such as
 * "[NC 50]" is one unit so it never splits across two lines.
 */
function units(text: string): string[] {
  return plainText(text).match(/(?:\[[^\]]*\]|\S)+/g) ?? [];
}

/**
 * Break text into lines no wider than `width`. Lines break at spaces; a unit wider than a line
 * on its own is cut where it runs out, as the sheet does for a long serial in a narrow cell.
 */
export function wrapText(text: string, bold: boolean, size: number, width: number): string[] {
  const lines: string[] = [];
  let line = "";
  const place = (unit: string): void => {
    const joined = line === "" ? unit : `${line} ${unit}`;
    if (textWidth(joined, bold, size) <= width) {
      line = joined;
      return;
    }
    if (line !== "") lines.push(line);
    line = "";
    if (unit.includes(" ") && textWidth(unit, bold, size) > width) {
      for (const word of unit.split(" ")) place(word);
      return;
    }
    let rest = unit;
    while (textWidth(rest, bold, size) > width && rest.length > 1) {
      let cut = rest.length - 1;
      while (cut > 1 && textWidth(rest.slice(0, cut), bold, size) > width) cut -= 1;
      lines.push(rest.slice(0, cut));
      rest = rest.slice(cut);
    }
    line = rest;
  };
  for (const unit of units(text)) place(unit);
  if (line !== "") lines.push(line);
  return lines;
}
