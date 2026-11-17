#!/usr/bin/env node
// Segment pictures (RenderCorridorExample) every N metres from surface, read straight from a
// workbook file. The workbook bytes are only read; nothing is written back to it.
//
//   pnpm exec tsx src/packages/corridor-3d/scripts/renderSegments.ts --workbook "Slide Sheet - 35783.xlsm" --step 500 --out docs/corridor_segments/500m

import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { createCanvas, GlobalFonts } from "@napi-rs/canvas";
import { readWorkbook } from "../../xlsx-bridge/src/readWorkbook";
import { SheetBook } from "../../workbook-commands/src/sheetBook";
import type { SnapshotBook } from "../../workbook-commands/src/cells";
import { corridorSize, paintCorridor } from "../src/corridorDraw";
import { CorridorScene } from "../src/corridorScene";
import type { PaintContext, Theme } from "../src/corridorPaint";

interface Args {
  workbook: string;
  step: number;
  out: string;
  scale: number;
  theme: Theme;
}

function parseArgs(argv: string[]): Args {
  const get = (name: string, fallback: string) => {
    const index = argv.indexOf(`--${name}`);
    return index >= 0 && argv[index + 1] ? argv[index + 1]! : fallback;
  };
  const theme = get("theme", "dark");
  if (theme !== "dark" && theme !== "paper") throw new Error(`--theme must be dark or paper, not ${theme}`);
  return {
    workbook: path.resolve(get("workbook", "Slide Sheet - Demo.xlsm")),
    step: Number(get("step", "500")),
    out: path.resolve(get("out", "docs/corridor_segments/500m")),
    scale: Number(get("scale", "2")),
    theme,
  };
}

const rule = (ch: string) => ch.repeat(60);
const clock = (ms: number) => new Date(ms).toISOString().slice(11, 19);

function registerFonts(): void {
  const fonts = process.env.WINDIR ? path.join(process.env.WINDIR, "Fonts") : "";
  for (const [file, family] of [["calibri.ttf", "Calibri"], ["calibrib.ttf", "Calibri"]] as const) {
    const full = path.join(fonts, file);
    if (fonts && existsSync(full)) GlobalFonts.registerFromPath(full, family);
  }
}

async function main(): Promise<void> {
  const args = parseArgs(process.argv.slice(2));
  const started = Date.now();
  console.log(rule("="));
  console.log(` CORRIDOR SEGMENTS | ${path.basename(args.workbook)} | every ${args.step} m`);
  console.log(` Timestamp: ${new Date().toISOString()}`);
  console.log(rule("="));
  console.log("Steps:\n [1/3] Validate inputs\n [2/3] Read workbook (read-only)\n [3/3] Render segments");
  console.log(rule("-"));

  console.log("[STEP 1/3] Validate inputs");
  if (!existsSync(args.workbook)) throw new Error(`Workbook not found: ${args.workbook}`);
  if (!(args.step > 0)) throw new Error(`--step must be positive, not ${args.step}`);
  if (!(args.scale > 0)) throw new Error(`--scale must be positive, not ${args.scale}`);
  mkdirSync(args.out, { recursive: true });
  registerFonts();
  console.log(`  Output dir: ${args.out}\n  Scale     : ${args.scale} px/pt\n  Theme     : ${args.theme}`);
  console.log(rule("-"));

  console.log("[STEP 2/3] Read workbook (read-only)");
  const t2 = Date.now();
  const mapped = await readWorkbook(new Uint8Array(readFileSync(args.workbook)));
  const book = new SheetBook(mapped.workbook as unknown as SnapshotBook);
  const live = CorridorScene.build(book);
  if (!live) throw new Error("No plan (_OC_Survey) or fewer than two good surveys on the Slidesheet.");
  const end = live.svyMd;
  console.log(`  Sheets    : ${mapped.importLog.sheetCount}\n  Last survey: ${end.toFixed(2)} m (inc ${live.svyInc.toFixed(2)})\n  Duration  : ${clock(Date.now() - t2)}`);
  console.log(rule("-"));

  console.log("[STEP 3/3] Render segments");
  const windows: Array<[number, number]> = [];
  for (let from = 0; from < end - 1; from += args.step) windows.push([from, Math.min(from + args.step, Math.ceil(end))]);
  const written: string[] = [];
  const skipped: string[] = [];
  const t3 = Date.now();
  windows.forEach(([from, to], index) => {
    const name = `seg_${String(from).padStart(4, "0")}_${String(to).padStart(4, "0")}.png`;
    const scene = CorridorScene.build(book, { from, to });
    if (scene) {
      const size = corridorSize(scene);
      const canvas = createCanvas(Math.round(size.width * args.scale), Math.round(size.height * args.scale));
      const ctx = canvas.getContext("2d");
      ctx.scale(args.scale, args.scale);
      paintCorridor(ctx as unknown as PaintContext, scene, args.theme);
      writeFileSync(path.join(args.out, name), canvas.toBuffer("image/png"));
      written.push(`${name}  ${scene.section.padEnd(8)} ${scene.sCount} stations`);
    } else {
      skipped.push(`${name}  fewer than two surveys in the window`);
    }
    const done = index + 1;
    const rate = done / Math.max(0.001, (Date.now() - t3) / 1000);
    process.stdout.write(`\r[STEP 3/3] Rendering [ ${done}/${windows.length} ] [ ${String(Math.floor(done / windows.length * 100)).padStart(3)}% ] ${rate.toFixed(1)} pictures/s`);
  });
  process.stdout.write("\n");
  console.log(rule("="));
  console.log(`RUN COMPLETE | ${written.length} written, ${skipped.length} skipped | ${clock(Date.now() - started)}`);
  for (const line of written) console.log(`  ${line}`);
  for (const line of skipped) console.log(`  SKIPPED ${line}`);
  console.log(rule("="));
}

main().catch((error: unknown) => {
  console.log(`\n${rule("=")}\nRUN FAILED\nError: ${error instanceof Error ? error.message : String(error)}\n${rule("=")}`);
  process.exit(1);
});
