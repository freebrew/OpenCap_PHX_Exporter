// Drawing primitives of MDL_CorridorImage (Rect / Ln / Dot / PolyLine / PolyFill / Tx) on a
// 2D canvas. Coordinates are points on the workbook's 1023-wide grid; the caller scales the
// context. Colours are the call-site hex names, mapped through the same DarkTone / Print*
// tables the workbook uses, so DARK_THEME pictures match the mail.

export type Dash = "solid" | "dash" | "dot";
export type Anchor = "start" | "middle" | "end";
export type Theme = "dark" | "paper";

/** The subset of CanvasRenderingContext2D the picture uses (browser and @napi-rs/canvas). */
export interface PaintContext {
  fillStyle: unknown;
  strokeStyle: unknown;
  lineWidth: number;
  lineCap: string;
  lineJoin: string;
  font: string;
  textAlign: string;
  textBaseline: string;
  beginPath(): void;
  closePath(): void;
  moveTo(x: number, y: number): void;
  lineTo(x: number, y: number): void;
  arc(x: number, y: number, r: number, a0: number, a1: number): void;
  rect(x: number, y: number, w: number, h: number): void;
  fill(): void;
  stroke(): void;
  setLineDash(segments: number[]): void;
  fillText(text: string, x: number, y: number): void;
  measureText(text: string): { width: number };
}

export const CANVAS_W = 1023;
const FONT = "Calibri, Carlito, 'Segoe UI', Arial, sans-serif";

const DARK_FILL: Record<string, string | null> = {
  FFFFFF: "262626", "1A1A1A": "262626", "222222": "262626", "2A2A2A": "262626", "242424": "262626", "252825": "262626", "111111": "262626",
  E8EBEE: "333333", "3D2458": "B08AF0", "6B2020": "FF6B6B", "1A3A5C": "6FA8DC",
  "1E3A2F": null, "1A3328": null, "1A3348": null,
};

const DARK_LINE: Record<string, string> = {
  "1A1A1A": "262626", "333333": "666666", "3A3A3A": "707070", "444444": "808080", "4A4A4A": "9A9A9A", "555555": "A8A8A8",
  "6A6A6A": "A0A0A0", "777777": "A0A0A0", "888888": "A0A0A0", "5C656C": "9AA5AD", B4BCC2: "5A6068", "7A848C": "9AA3AA",
  "008C95": "2EC4CF", C0392B: "FF6A5A", "6B2020": "FF6B6B", "1A3A5C": "6FA8DC", "2C3338": "E6E6E6", "6B4F2A": "C9A66B",
  "1B4D2E": "4CAF7A", "8B0000": "FF5555", "6B4A00": "E09A3D",
};

const DARK_TEXT: Record<string, string> = {
  A8A8A8: "BFBFBF", "7A7A7A": "A6A6A6", "6B7A85": "A6B0B8", "5C6770": "B8C0C6", "008C95": "2EC4CF", "6B4F2A": "C9A66B",
  "6B2020": "FF6B6B", "1A3A5C": "6FA8DC", "7A848C": "9AA3AA", "5C656C": "9AA5AD", "1B4D2E": "4CAF7A", "8B0000": "FF5555",
  "6B4A00": "E09A3D", C0392B: "FF6A5A", "2C3338": "F2F2F2", "1A1A1A": "F2F2F2", "4A4A4A": "BFBFBF", "555555": "BFBFBF",
};

const PAPER_LINE: Record<string, string> = {
  FFFFFF: "1A1A1A", A8CBB8: "777777", B6DBC8: "777777", "9EC9B0": "777777", "3A3A3A": "B0B0B0", "333333": "B0B0B0", "444444": "B0B0B0",
  "5BA3D9": "1A3A5C", "4CAF7A": "1B4D2E", "3D6B54": "1B4D2E", FF5555: "8B0000", E09A3D: "6B4A00",
};

const PAPER_TEXT: Record<string, string> = {
  FFFFFF: "1A1A1A", A8A8A8: "4A4A4A", "7A7A7A": "555555", "8AAB9A": "4A4A4A", "4CAF7A": "1B4D2E", "5BA3D9": "1A3A5C",
  FF5555: "8B0000", E09A3D: "6B4A00", B6DBC8: "555555",
};

function paperFill(hex: string): string | null {
  if (["1A1A1A", "222222", "2A2A2A", "242424", "252825"].includes(hex)) return "FFFFFF";
  if (["1E3A2F", "1A3328", "1A3348"].includes(hex)) return null;
  if (hex === "B6DBC8") return "888888";
  return hex;
}

function css(hex: string | null): string | null {
  return hex == null ? null : `#${hex}`;
}

function dashOf(dash: Dash, weight: number): number[] {
  switch (dash) {
    case "solid":
      return [];
    case "dash":
      return [4 * weight, 3 * weight];
    case "dot":
      return [weight, weight];
    default: {
      const never: never = dash;
      return never;
    }
  }
}

export class Painter {
  /** Applied inside tx, like mTypeScale. */
  typeScale = 1;

  constructor(readonly ctx: PaintContext, readonly theme: Theme = "dark") {}

  private fillColor(hex: string | null): string | null {
    if (hex == null) return null;
    const key = hex.toUpperCase();
    if (this.theme === "dark") return css(key in DARK_FILL ? DARK_FILL[key]! : key);
    return css(paperFill(key));
  }

  private lineColor(hex: string | null): string | null {
    if (hex == null) return null;
    const key = hex.toUpperCase();
    if (this.theme === "dark") return css(DARK_LINE[key] ?? key);
    return css(PAPER_LINE[key] ?? key);
  }

  private textColor(hex: string): string {
    const key = hex.toUpperCase();
    if (this.theme === "dark") return `#${DARK_TEXT[key] ?? key}`;
    return `#${PAPER_TEXT[key] ?? key}`;
  }

  private strokePath(color: string | null, weight: number, dash: Dash): void {
    const stroke = this.lineColor(color);
    if (stroke == null) return;
    this.ctx.strokeStyle = stroke;
    this.ctx.lineWidth = weight;
    this.ctx.lineCap = dash === "solid" ? "round" : "butt";
    this.ctx.lineJoin = "round";
    this.ctx.setLineDash(dashOf(dash, weight));
    this.ctx.stroke();
    this.ctx.setLineDash([]);
  }

  private fillPath(color: string | null): void {
    const fill = this.fillColor(color);
    if (fill == null) return;
    this.ctx.fillStyle = fill;
    this.ctx.fill();
  }

  rect(l: number, t: number, w: number, h: number, fill: string | null, line: string | null, weight: number): void {
    this.ctx.beginPath();
    this.ctx.rect(l, t, Math.max(w, 0.1), Math.max(h, 0.1));
    this.fillPath(fill);
    this.strokePath(line, weight, "solid");
  }

  ln(x1: number, y1: number, x2: number, y2: number, color: string, weight: number, dash: Dash): void {
    this.ctx.beginPath();
    this.ctx.moveTo(x1, y1);
    this.ctx.lineTo(x2, y2);
    this.strokePath(color, weight, dash);
  }

  dot(cx: number, cy: number, r: number, fill: string | null, line: string | null, weight: number): void {
    this.ctx.beginPath();
    this.ctx.arc(cx, cy, r, 0, Math.PI * 2);
    this.fillPath(fill);
    if (weight > 0) this.strokePath(line, weight, "solid");
  }

  polyLine(xs: number[], ys: number[], n: number, color: string, weight: number, dash: Dash): void {
    if (n < 2) return;
    this.ctx.beginPath();
    this.ctx.moveTo(xs[0]!, ys[0]!);
    for (let i = 1; i < n; i += 1) this.ctx.lineTo(xs[i]!, ys[i]!);
    this.strokePath(color, weight, dash);
  }

  polyFill(xs: number[], ys: number[], n: number, fill: string | null, line: string | null, weight: number): void {
    if (n < 3) return;
    this.ctx.beginPath();
    this.ctx.moveTo(xs[0]!, ys[0]!);
    for (let i = 1; i < n; i += 1) this.ctx.lineTo(xs[i]!, ys[i]!);
    this.ctx.closePath();
    this.fillPath(fill);
    this.strokePath(line, weight, "solid");
  }

  /** Width of a run at the live type scale (MeasureTx). */
  measure(text: string, size: number, bold: boolean): number {
    if (text === "") return 0;
    this.ctx.font = `${bold ? "bold " : ""}${size * this.typeScale}px ${FONT}`;
    return this.ctx.measureText(text).width;
  }

  /**
   * Tx: x is the anchor and y sits where the workbook's text box puts the run (box top at
   * y - 1.25 sz, Calibri baseline about 0.95 sz below the top).
   */
  tx(x: number, y: number, text: string, size: number, color: string, anchor: Anchor, bold = false): void {
    const sz = size * this.typeScale;
    this.ctx.font = `${bold ? "bold " : ""}${sz}px ${FONT}`;
    this.ctx.fillStyle = this.textColor(color);
    this.ctx.textBaseline = "alphabetic";
    this.ctx.textAlign = anchor === "start" ? "left" : anchor === "end" ? "right" : "center";
    this.ctx.fillText(text, x, y - 0.3 * sz);
    this.ctx.textAlign = "left";
  }

  /** TxWallLabel: name, TVD, and lithology stepped along the wall line. */
  wallLabel(x1: number, y1: number, x2: number, y2: number, name: string, tvdLabel: string, lithology: string, size: number, color: string): void {
    const dx = x2 - x1;
    const dy = y2 - y1;
    this.tx(x1 + dx * 0.02, y1 + dy * 0.02, name, size, color, "start", true);
    this.tx(x1 + dx * 0.35, y1 + dy * 0.35, tvdLabel, size, color, "start");
    if (lithology.trim() !== "") this.tx(x1 + dx * 0.68, y1 + dy * 0.68, lithology, size, color, "start");
  }
}
