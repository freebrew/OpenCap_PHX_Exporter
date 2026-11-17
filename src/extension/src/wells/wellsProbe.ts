import type { SnapshotBook } from "../../../packages/workbook-commands/src/cells";
import { SheetBook } from "../../../packages/workbook-commands/src/sheetBook";
import { GOOD_SURVEY_FILL } from "../../../packages/workbook-commands/src/slidesheet";
import { shownOffsets } from "../../../packages/wells-3d/src/wellFilter";
import { buildWellScene } from "../../../packages/wells-3d/src/wellScene";
import type { HoverPick } from "../../../packages/wells-3d/src/wellPick";
import type { ProbeResult, WellsIn, WellsOut } from "./wellsMessages";

interface ReadyResult {
  backend: "webgpu" | "webgl2";
}

export interface WellsProbeResult {
  forced: ProbeResult;
  open: ReadyResult;
  hover: HoverPick | null;
  /** The same hover after the filter hides every offset well. Its SF is then null. */
  hiddenSf: number | null;
  sample: { center: [number, number, number]; corner: [number, number, number] };
  bench: { maxMs: number; frames: number; mainBusyMs: number };
}

interface ProbeWindow extends Window {
  __phxWells?: Promise<WellsProbeResult>;
}

function book(): SheetBook {
  const sheets = ["_OC_Survey", "_OC_AC", "_OC_AC_Detail", "_OC_Formations", "Data", "Slidesheet"];
  const out: Record<string, { id: string; name: string; cellData: Record<string, never> }> = {};
  sheets.forEach((name, index) => {
    out[`s${index}`] = { id: `s${index}`, name, cellData: {} };
  });
  return new SheetBook({ sheets: out, styles: {} } as SnapshotBook, []);
}

/** Invented plan and one concern well. The look-at point sits on the plan, and the printed SF there is 1.2. */
function fixtureScene() {
  const sheet = book();
  sheet.setValue("_OC_Survey", 3, 1, 0);
  sheet.setValue("_OC_Survey", 3, 4, 0);
  sheet.setValue("_OC_Survey", 3, 5, 0);
  sheet.setValue("_OC_Survey", 3, 6, 0);
  sheet.setValue("_OC_Survey", 4, 1, 1000);
  sheet.setValue("_OC_Survey", 4, 4, 400);
  sheet.setValue("_OC_Survey", 4, 5, 100);
  sheet.setValue("_OC_Survey", 4, 6, 500);
  sheet.setValue("_OC_Formations", 3, 1, 500);
  sheet.setValue("_OC_Formations", 3, 2, 200);
  sheet.setValue("_OC_Formations", 3, 3, "Belly River");
  sheet.setValue("_OC_Formations", 3, 4, "sandstone");
  sheet.setValue("Data", 10, 2, "AC Info & Concerns");
  sheet.setValue("Data", 14, 2, "Motors On Location");
  sheet.setValue("_OC_AC", 3, 1, "103/09-10-011-12W5/0 - East Offset");
  sheet.setValue("_OC_AC", 3, 2, 100);
  sheet.setValue("_OC_AC", 3, 3, 15);
  sheet.setValue("_OC_AC", 3, 4, 1.1);
  sheet.setValue("_OC_AC", 4, 1, "100/01-02-003-04W5/0 - North Offset");
  sheet.setValue("_OC_AC", 4, 2, 500);
  sheet.setValue("_OC_AC", 4, 3, 20);
  sheet.setValue("_OC_AC", 4, 4, 1.2);
  sheet.setValue("_OC_AC", 5, 1, "102/05-06-007-08W5/2 - South Offset");
  sheet.setValue("_OC_AC", 5, 2, 800);
  sheet.setValue("_OC_AC", 5, 3, 40);
  sheet.setValue("_OC_AC", 5, 4, 4.5);
  sheet.setValue("_OC_AC_Detail", 3, 1, "100/01-02-003-04W5/0 - North Offset");
  sheet.setValue("_OC_AC_Detail", 3, 2, 500);
  sheet.setValue("_OC_AC_Detail", 3, 3, 490);
  sheet.setValue("_OC_AC_Detail", 3, 4, 220);
  sheet.setValue("_OC_AC_Detail", 3, 5, 50);
  sheet.setValue("_OC_AC_Detail", 3, 6, 250);
  sheet.setValue("_OC_AC_Detail", 3, 7, 20);
  sheet.setValue("_OC_AC_Detail", 3, 8, 1.2);
  sheet.setValue("_OC_AC_Detail", 4, 1, "102/05-06-007-08W5/2 - South Offset");
  sheet.setValue("_OC_AC_Detail", 4, 2, 800);
  sheet.setValue("_OC_AC_Detail", 4, 3, 790);
  sheet.setValue("_OC_AC_Detail", 4, 4, 360);
  sheet.setValue("_OC_AC_Detail", 4, 5, 80);
  sheet.setValue("_OC_AC_Detail", 4, 6, 400);
  sheet.setValue("_OC_AC_Detail", 4, 7, 40);
  sheet.setValue("_OC_AC_Detail", 4, 8, 4.5);
  for (const row of [13, 14]) {
    sheet.setValue("Slidesheet", row, 5, row === 13 ? 100 : 200);
    sheet.setValue("Slidesheet", row, 6, 0);
    sheet.setValue("Slidesheet", row, 7, 0);
    sheet.setFill("Slidesheet", { row, col: 6, row2: row, col2: 6 }, GOOD_SURVEY_FILL);
    sheet.setFill("Slidesheet", { row, col: 7, row2: row, col2: 7 }, GOOD_SURVEY_FILL);
  }
  return buildWellScene(sheet);
}

function openWorker(canvas: HTMLCanvasElement, forceWebGL: boolean): {
  ready: Promise<ReadyResult>;
  send: (message: WellsIn) => void;
  ask: (message: WellsIn) => Promise<WellsOut>;
} {
  const offscreen = canvas.transferControlToOffscreen();
  const worker = new Worker(new URL("./wells.worker.ts", import.meta.url), { type: "module" });
  let waiter: ((message: WellsOut) => void) | null = null;
  let fail: (error: Error) => void = () => {};
  const ready = new Promise<ReadyResult>((resolve, reject) => {
    fail = reject;
    worker.onmessage = (event: MessageEvent<WellsOut>) => {
      const message = event.data;
      if (message.type === "error") {
        fail(new Error(message.message));
        return;
      }
      if (message.type === "ready") {
        resolve({ backend: message.backend });
        return;
      }
      waiter?.(message);
      waiter = null;
    };
    worker.onerror = (event) => fail(new Error(event.message));
  });
  worker.postMessage({
    type: "init",
    canvas: offscreen,
    width: canvas.clientWidth || canvas.width,
    height: canvas.clientHeight || canvas.height,
    dpr: 1,
    forceWebGL,
  }, [offscreen]);
  return {
    ready,
    send(message: WellsIn) {
      worker.postMessage(message);
    },
    ask(message: WellsIn) {
      return new Promise((resolve, reject) => {
        fail = reject;
        waiter = resolve;
        worker.postMessage(message);
      });
    },
  };
}

async function run(): Promise<WellsProbeResult> {
  const glCanvas = document.querySelector<HTMLCanvasElement>("#gl");
  const gpuCanvas = document.querySelector<HTMLCanvasElement>("#gpu");
  if (!glCanvas || !gpuCanvas) throw new Error("Probe canvases are missing");
  const forced = openWorker(glCanvas, true);
  const open = openWorker(gpuCanvas, false);
  const [forcedReady, openReady] = await Promise.all([forced.ready, open.ready]);
  const probe = await forced.ask({ type: "probe" });
  if (probe.type !== "probe") throw new Error(`Probe replied ${probe.type}`);
  const scene = fixtureScene();
  if (!scene) throw new Error("Fixture scene is empty");
  const shown = shownOffsets(scene.offsets, { allOffsets: false, site: "all", maxSf: null, maxC2c: null });
  forced.send({ type: "options", drilled: false, shown });
  forced.send({ type: "scene", scene });
  const sampleMessage = await forced.ask({ type: "sample" });
  if (sampleMessage.type !== "sample") throw new Error(`Sample replied ${sampleMessage.type}`);
  const width = glCanvas.clientWidth || glCanvas.width;
  const height = glCanvas.clientHeight || glCanvas.height;
  const hoverMessage = await forced.ask({ type: "hover", x: width / 2, y: height / 2 });
  if (hoverMessage.type !== "hover") throw new Error(`Hover replied ${hoverMessage.type}`);
  forced.send({ type: "options", drilled: false, shown: [] });
  const hiddenMessage = await forced.ask({ type: "hover", x: width / 2, y: height / 2 });
  if (hiddenMessage.type !== "hover") throw new Error(`Hidden hover replied ${hiddenMessage.type}`);
  const started = performance.now();
  const benchPromise = forced.ask({ type: "bench", frames: 30 });
  while (performance.now() - started < 200) {
    // The page thread stays busy on purpose. The frames are drawn on the worker.
  }
  const benchMessage = await benchPromise;
  if (benchMessage.type !== "bench") throw new Error(`Bench replied ${benchMessage.type}`);
  return {
    forced: probe,
    open: openReady,
    hover: hoverMessage.hit,
    hiddenSf: hiddenMessage.hit?.sf ?? null,
    sample: { center: sampleMessage.center, corner: sampleMessage.corner },
    bench: { maxMs: benchMessage.maxMs, frames: benchMessage.frames, mainBusyMs: performance.now() - started },
  };
}

(window as ProbeWindow).__phxWells = run();
