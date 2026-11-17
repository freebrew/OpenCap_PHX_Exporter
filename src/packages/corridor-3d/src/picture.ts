import type { SheetBook } from "../../workbook-commands/src/sheetBook";
import { corridorSize, paintCorridor } from "./corridorDraw";
import type { PaintContext, Theme } from "./corridorPaint";
import { CorridorScene, type Section } from "./corridorScene";

export interface CorridorPicture {
  section: Section;
  /** "1:1 shaft" or "MD tunnel" plus the window, for captions and status lines. */
  caption: string;
  width: number;
  height: number;
}

export interface CorridorOptions {
  /** Pixels per point; 1 gives the 1023 px picture the mail is laid out for. */
  scale?: number;
  theme?: Theme;
  /** RenderCorridorExample window instead of the live report day. */
  example?: { from: number; to: number };
}

/**
 * MDL_CorridorImage.RenderCorridorPng(path, "corridor") into a canvas. Null when the
 * workbook has no plan or the window holds fewer than two good surveys.
 */
export function renderCorridor(canvas: HTMLCanvasElement, book: SheetBook, options: CorridorOptions = {}): CorridorPicture | null {
  const scene = CorridorScene.build(book, options.example);
  if (!scene) return null;
  const scale = options.scale ?? 1;
  const size = corridorSize(scene);
  canvas.width = Math.round(size.width * scale);
  canvas.height = Math.round(size.height * scale);
  const context = canvas.getContext("2d");
  if (!context) return null;
  context.setTransform(scale, 0, 0, scale, 0, 0);
  paintCorridor(context as unknown as PaintContext, scene, options.theme ?? "dark");
  const view = scene.shaft ? "1:1 shaft" : "MD tunnel";
  const from = Math.round(scene.isExample ? scene.exA : scene.md0);
  const to = Math.round(scene.isExample ? scene.exB : scene.md1);
  return { section: scene.section, caption: `${view} ${scene.section.toLowerCase()} · ${from} to ${to} m`, width: canvas.width, height: canvas.height };
}
