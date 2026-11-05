// A small PDF writer for the strap card: pages of drawing operators, the two standard Helvetica
// fonts, and PNG-predictor images. Bytes are assembled as bytes, so an image stream survives.

import type { CardArt, ImageArt, VectorItem, VectorOp } from "./cardSpec";
import { plainText } from "./helvetica";

export type Rgb = readonly [number, number, number];

const encoder = new TextEncoder();

function num(value: number): string {
  return String(Math.round(value * 100) / 100);
}

function color(value: Rgb): string {
  return value.map((part) => num(part)).join(" ");
}

function literal(text: string): string {
  return `(${plainText(text).replaceAll("\\", "\\\\").replaceAll("(", "\\(").replaceAll(")", "\\)")})`;
}

function fromBase64(text: string): Uint8Array {
  return Uint8Array.from(atob(text), (character) => character.charCodeAt(0));
}

function hex(bytes: Uint8Array): string {
  return Array.from(bytes, (byte) => byte.toString(16).padStart(2, "0")).join("");
}

/** Where layout units land on the page: layout (0,0) is the content's top-left corner. */
export interface PageFrame {
  /** Page points from the left edge to layout x = 0. */
  left: number;
  /** Page points from the top edge to layout y = 0. */
  top: number;
  /** Page points per layout unit. */
  scale: number;
  /** Page height in points. */
  height: number;
}

/** One page's drawing operators, in layout units (y grows downward). */
export class PageCanvas {
  private readonly ops: string[] = [];
  readonly images = new Set<string>();

  constructor(private readonly frame: PageFrame) {}

  private px(x: number): number {
    return this.frame.left + x * this.frame.scale;
  }

  private py(y: number): number {
    return this.frame.height - (this.frame.top + y * this.frame.scale);
  }

  fillRect(x: number, y: number, w: number, h: number, fill: Rgb): void {
    this.ops.push(`${color(fill)} rg ${num(this.px(x))} ${num(this.py(y + h))} ${num(w * this.frame.scale)} ${num(h * this.frame.scale)} re f`);
  }

  strokeRect(x: number, y: number, w: number, h: number, width: number): void {
    this.ops.push(`0 0 0 RG ${num(width * this.frame.scale)} w ${num(this.px(x))} ${num(this.py(y + h))} ${num(w * this.frame.scale)} ${num(h * this.frame.scale)} re S`);
  }

  line(x1: number, y1: number, x2: number, y2: number, width: number): void {
    this.ops.push(`0 0 0 RG ${num(width * this.frame.scale)} w ${num(this.px(x1))} ${num(this.py(y1))} m ${num(this.px(x2))} ${num(this.py(y2))} l S`);
  }

  /** Text with its baseline at `baseline`, `size` in layout units. */
  text(text: string, x: number, baseline: number, size: number, bold: boolean, fill: Rgb): void {
    this.ops.push(`BT ${color(fill)} rg /${bold ? "F2" : "F1"} ${num(size * this.frame.scale)} Tf ${num(this.px(x))} ${num(this.py(baseline))} Td ${literal(text)} Tj ET`);
  }

  /** Whatever `draw` adds is cut off outside the rectangle. */
  clipped(x: number, y: number, w: number, h: number, draw: () => void): void {
    this.ops.push(`q ${num(this.px(x))} ${num(this.py(y + h))} ${num(w * this.frame.scale)} ${num(h * this.frame.scale)} re W n`);
    draw();
    this.ops.push("Q");
  }

  image(name: string, x: number, y: number, w: number, h: number): void {
    this.images.add(name);
    this.ops.push(`q ${num(w * this.frame.scale)} 0 0 ${num(h * this.frame.scale)} ${num(this.px(x))} ${num(this.py(y + h))} cm /${name} Do Q`);
  }

  /** Paint one drawing. `sx` and `sy` are layout units per art unit. */
  paintItems(items: VectorItem[], ox: number, oy: number, sx: number, sy: number): void {
    for (const item of items) {
      if (item.kind === "line") {
        const dash = dashOf(item.dash, sx * this.frame.scale);
        this.ops.push(
          `q ${rgbOf(item.stroke)} RG ${num(item.width * sx * this.frame.scale)} w ${dash}` +
            `${this.point(ox + item.x1 * sx, oy + item.y1 * sy)} m ` +
            `${this.point(ox + item.x2 * sx, oy + item.y2 * sy)} l S Q`,
        );
        continue;
      }
      const fill = item.fill && item.fill !== "none" ? rgbOf(item.fill) : "";
      const stroke = item.stroke && item.stroke !== "none" ? rgbOf(item.stroke) : "";
      const width = item.width ? `${num(item.width * sx * this.frame.scale)} w ` : "";
      const dash = dashOf(item.dash, sx * this.frame.scale);
      const paint = fill && stroke ? (item.evenodd ? "B*" : "B") : fill ? (item.evenodd ? "f*" : "f") : "S";
      this.ops.push(`q ${fill ? `${fill} rg ` : ""}${stroke ? `${stroke} RG ` : ""}${width}${dash}${this.pathOps(item.ops, ox, oy, sx, sy)} ${paint} Q`);
    }
  }

  private point(x: number, y: number): string {
    return `${num(this.px(x))} ${num(this.py(y))}`;
  }

  private pathOps(ops: VectorOp[], ox: number, oy: number, sx: number, sy: number): string {
    return ops
      .map((op) => {
        if (op.op === "Z") return "h";
        if (op.op === "C") {
          return `${this.point(ox + op.x1 * sx, oy + op.y1 * sy)} ${this.point(ox + op.x2 * sx, oy + op.y2 * sy)} ${this.point(ox + op.x * sx, oy + op.y * sy)} c`;
        }
        return `${this.point(ox + op.x * sx, oy + op.y * sy)} ${op.op === "M" ? "m" : "l"}`;
      })
      .join(" ");
  }

  content(): string {
    return this.ops.join("\n");
  }
}

interface Placed {
  name: string;
  art: CardArt;
}

/** The document: pages added in order, each drawing image shared by every page that uses it. */
export class PdfWriter {
  private readonly placed = new Map<CardArt, Placed>();
  private readonly pages: PageCanvas[] = [];

  /** The resource name of a drawing, registering it the first time it is used. */
  imageName(art: CardArt): string {
    if (art.kind !== "image") throw new Error("vector drawings are painted, not embedded");
    let known = this.placed.get(art);
    if (!known) {
      known = { name: `Im${this.placed.size}`, art };
      this.placed.set(art, known);
    }
    return known.name;
  }

  addPage(page: PageCanvas): void {
    this.pages.push(page);
  }

  build(): Uint8Array {
    const placed = [...this.placed.values()];
    const imageIds = new Map(placed.map((item, index) => [item.name, 5 + index]));
    const firstPage = 5 + placed.length;
    const pageIds = this.pages.map((_, index) => firstPage + index * 2 + 1);

    const objects: Uint8Array[][] = [
      [encoder.encode("<< /Type /Catalog /Pages 2 0 R >>")],
      [encoder.encode(`<< /Type /Pages /Count ${this.pages.length} /Kids [${pageIds.map((id) => `${id} 0 R`).join(" ")}] >>`)],
      [encoder.encode("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>")],
      [encoder.encode("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>")],
    ];
    for (const item of placed) {
      if (item.art.kind !== "image") throw new Error("vector art has no image object");
      objects.push(imageObject(item.art));
    }
    this.pages.forEach((page, index) => {
      const contentId = firstPage + index * 2;
      const stream = encoder.encode(page.content());
      objects.push([encoder.encode(`<< /Length ${stream.length} >>\nstream\n`), stream, encoder.encode("\nendstream")]);
      const used = [...page.images].map((name) => `/${name} ${imageIds.get(name)} 0 R`).join(" ");
      objects.push([
        encoder.encode(
          `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 3 0 R /F2 4 0 R >>${used ? ` /XObject << ${used} >>` : ""} >> /Contents ${contentId} 0 R >>`,
        ),
      ]);
    });

    const chunks: Uint8Array[] = [new Uint8Array([...encoder.encode("%PDF-1.4\n%"), 0xe2, 0xe3, 0xcf, 0xd3, 0x0a])];
    let length = chunks[0]?.length ?? 0;
    const offsets: number[] = [];
    objects.forEach((parts, index) => {
      offsets.push(length);
      const open = encoder.encode(`${index + 1} 0 obj\n`);
      const close = encoder.encode("\nendobj\n");
      for (const part of [open, ...parts, close]) {
        chunks.push(part);
        length += part.length;
      }
    });
    const xref =
      `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n` +
      offsets.map((offset) => `${String(offset).padStart(10, "0")} 00000 n \n`).join("") +
      `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${length}\n%%EOF`;
    chunks.push(encoder.encode(xref));

    const out = new Uint8Array(chunks.reduce((sum, chunk) => sum + chunk.length, 0));
    let at = 0;
    for (const chunk of chunks) {
      out.set(chunk, at);
      at += chunk.length;
    }
    return out;
  }
}

/** The PNG's own deflate stream, told to undo PNG filtering with a predictor. */
function rgbOf(hex: string): string {
  const value = Number.parseInt(hex.slice(1), 16);
  const red = ((value >> 16) & 255) / 255;
  const green = ((value >> 8) & 255) / 255;
  const blue = (value & 255) / 255;
  return color([red, green, blue]);
}

function dashOf(dash: number[] | undefined, sx: number): string {
  if (!dash || dash.length === 0) return "";
  return `[${dash.map((part) => num(part * sx)).join(" ")}] 0 d `;
}

function imageObject(art: ImageArt): Uint8Array[] {
  const data = fromBase64(art.data);
  const space = art.indexed ? `[/Indexed /DeviceRGB ${fromBase64(art.palette).length / 3 - 1} <${hex(fromBase64(art.palette))}>]` : "/DeviceRGB";
  const colors = art.indexed ? 1 : 3;
  return [
    encoder.encode(
      `<< /Type /XObject /Subtype /Image /Width ${art.w} /Height ${art.h} /ColorSpace ${space} /BitsPerComponent 8 ` +
        `/Filter /FlateDecode /DecodeParms << /Predictor 15 /Colors ${colors} /BitsPerComponent 8 /Columns ${art.w} >> /Length ${data.length} >>\nstream\n`,
    ),
    data,
    encoder.encode("\nendstream"),
  ];
}
