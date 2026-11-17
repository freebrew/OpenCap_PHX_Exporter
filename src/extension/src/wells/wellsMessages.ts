import type { HoverPick } from "../../../packages/wells-3d/src/wellPick";
import type { WellScene } from "../../../packages/wells-3d/src/wellScene";

export type WellsBackend = "webgpu" | "webgl2";

export type WellsIn =
  | { type: "init"; canvas: OffscreenCanvas; width: number; height: number; dpr: number; forceWebGL: boolean }
  | { type: "resize"; width: number; height: number; dpr: number }
  | { type: "probe" }
  | { type: "scene"; scene: WellScene | null }
  | { type: "pointer"; phase: "down" | "move" | "up" | "wheel"; x: number; y: number; button: number; shift: boolean; deltaY: number }
  | { type: "hover"; x: number; y: number }
  | { type: "options"; drilled: boolean; shown: string[] }
  | { type: "reset" }
  | { type: "bench"; frames: number }
  | { type: "sample" };

export interface ProbeResult {
  type: "probe";
  backend: WellsBackend;
  wide: "line" | "tube";
  center: [number, number, number];
  shoulder: [number, number, number];
}

export type WellsOut =
  | { type: "ready"; backend: WellsBackend }
  | ProbeResult
  | { type: "hover"; hit: HoverPick | null }
  | { type: "bench"; maxMs: number; frames: number }
  | { type: "sample"; center: [number, number, number]; corner: [number, number, number] }
  | { type: "error"; message: string };
