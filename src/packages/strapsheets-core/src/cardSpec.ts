// The shapes the generated card tables (cards.gen.ts, art.gen.ts) fill in.

/** One printed card layout: the cells the Sheets print engine stamps for a tool family. */
export interface PrintCard {
  /** The print tab, e.g. "Motors". */
  name: string;
  /** The entry family it prints. Several tabs can print one family (Subs). */
  which: string;
  /** The big title under the cards. */
  title: string;
  /** Columns in one card. */
  width: number;
  /** Cards side by side in one band: 1, or 2 for the short sub cards. */
  positions: number;
  /** Columns from one card to the next in a band. */
  posShift: number;
  /** Height in sheet pixels of each of the five rows of a card. */
  rowHeights: number[];
  /** [row, col] of the serial cell inside a card. */
  serialCell: [number, number];
  /** Form Responses column holding the sub type, when a family splits across tabs. */
  typeCol: number | null;
  /** The sub types this tab prints. */
  typeValues: string[] | null;
  /** [row, col, text]: a label cell. */
  statics: [number, number, string][];
  /** [row, col, column] or [row, col, column, column]: a value cell, the columns joined by a space. */
  fields: ([number, number, number] | [number, number, number, number])[];
}

/** A tool drawing as the pieces a PDF image needs; the PNG is never decoded. */
export interface CardArt {
  w: number;
  h: number;
  /** Palette PNG (8-bit indexed) when true, else 8-bit RGB. */
  indexed: boolean;
  /** Base64 of the RGB palette, "" for RGB. */
  palette: string;
  /** Base64 of the PNG's joined IDAT stream. */
  data: string;
}
