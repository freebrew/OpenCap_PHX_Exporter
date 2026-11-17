import { WellsRenderer } from "./wellsRenderer";
import type { WellsIn, WellsOut } from "./wellsMessages";

let gpu: WellsRenderer | null = null;
const waiting: WellsIn[] = [];

function post(message: WellsOut): void {
  self.postMessage(message);
}

function handle(message: WellsIn): void {
  if (!gpu || message.type === "init") return;
  switch (message.type) {
    case "resize":
      gpu.resize(message.width, message.height, message.dpr);
      break;
    case "probe":
      post(gpu.probe());
      break;
    case "scene":
      gpu.setScene(message.scene);
      break;
    case "pointer":
      gpu.pointer(message.phase, message.x, message.y, message.button, message.shift, message.deltaY);
      break;
    case "hover":
      post({ type: "hover", hit: gpu.hover(message.x, message.y) });
      break;
    case "options":
      gpu.setOptions(message.drilled, message.shown);
      break;
    case "reset":
      gpu.reset();
      break;
    case "bench": {
      const result = gpu.bench(message.frames);
      post({ type: "bench", maxMs: result.maxMs, frames: result.frames });
      break;
    }
    case "sample": {
      const pixels = gpu.sample();
      post({ type: "sample", center: pixels.center, corner: pixels.corner });
      break;
    }
    default: {
      const never: never = message;
      post({ type: "error", message: `Unknown message ${String(never)}` });
    }
  }
}

self.onmessage = (event: MessageEvent<WellsIn>) => {
  const message = event.data;
  if (message.type === "init") {
    gpu = new WellsRenderer(message.canvas);
    gpu.init(message.forceWebGL, message.width, message.height, message.dpr).then((backend) => {
      post({ type: "ready", backend });
      for (const queued of waiting) {
        try {
          handle(queued);
        } catch (error) {
          post({ type: "error", message: error instanceof Error ? error.message : String(error) });
        }
      }
      waiting.length = 0;
    }).catch((error: unknown) => {
      post({ type: "error", message: error instanceof Error ? error.message : String(error) });
    });
    return;
  }
  if (!gpu) {
    waiting.push(message);
    return;
  }
  try {
    handle(message);
  } catch (error) {
    post({ type: "error", message: error instanceof Error ? error.message : String(error) });
  }
};
