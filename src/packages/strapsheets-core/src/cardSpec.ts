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
  /** [first column, column count] of each header cell merged across columns in a card's first label and value rows. */
  headerSpans?: [number, number][];
  /** The column where the yellow block of the second label row starts, when its labels are not all "A LEN"-style. */
  metricFrom?: number;
}

export type VectorOp =
  | { op: "M" | "L"; x: number; y: number }
  | { op: "C"; x1: number; y1: number; x2: number; y2: number; x: number; y: number }
  | { op: "Z" };

export type VectorItem =
  | {
      kind: "path";
      ops: VectorOp[];
      fill?: string;
      stroke?: string;
      width?: number;
      dash?: number[];
      evenodd?: boolean;
    }
  | {
      kind: "line";
      x1: number;
      y1: number;
      x2: number;
      y2: number;
      stroke: string;
      width: number;
      dash?: number[];
    };

/** A tool drawing as path operators. Diagram letters are outlines, not a PDF font. */
export interface VectorArt {
  kind: "vector";
  w: number;
  h: number;
  items: VectorItem[];
}

/** Kept so an older bitmap drawing can still be named. The shipped cards use vectors. */
export interface ImageArt {
  kind: "image";
  w: number;
  h: number;
  indexed: boolean;
  palette: string;
  data: string;
}

export type CardArt = VectorArt | ImageArt;
