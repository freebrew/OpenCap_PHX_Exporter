import * as THREE from "three";
import { Line2NodeMaterial, WebGPURenderer } from "three/webgpu";
import { Line2 } from "three/addons/lines/webgpu/Line2.js";
import { LineGeometry } from "three/addons/lines/LineGeometry.js";
import { dragOrbit, emptyBounds, fitOrbit, includeXyz, wheelOrbit, cameraFromOrbit, type Orbit } from "../../../packages/wells-3d/src/orbit";
import { shadowBoxOf } from "../../../packages/wells-3d/src/shadowBox";
import { pickHover, unpackPath, type PickPath } from "../../../packages/wells-3d/src/wellPick";
import type { OffsetPath, WellScene } from "../../../packages/wells-3d/src/wellScene";
import type { ProbeResult, WellsBackend } from "./wellsMessages";
import type { AnyObject } from "./wellsObject";
import { ShadowWalls } from "./wellsWalls";

const PLAN_COLOR = "#17324d";
const DRILLED_COLOR = "#0e7490";
const TARGET_COLOR = "#1b4d2e";
// Screen pixels. The plan and the drilled hole are the pair being compared, so they are drawn
// far heavier than the offset wells. A metre-sized pipe would swallow a close separation.
const PATH_WIDTH = 16;
const BACKGROUND = "#d5dee6";

interface Drawn {
  object: AnyObject;
  kind: "plan" | "offset" | "drilled";
  concern: boolean;
  name: string;
}

/** The polyline cut at its breaks, one piece per run. The pieces share the array. */
function runsOf(xyz: Float32Array, breaks: number[]): Float32Array[] {
  const out: Float32Array[] = [];
  let from = 0;
  for (const at of [...breaks, xyz.length / 3]) {
    out.push(xyz.subarray(from * 3, at * 3));
    from = at;
  }
  return out;
}

/** GPU picture of a well scene. Frames are drawn only when the camera, the hover or the scene changes. */
export class WellsRenderer {
  private renderer: {
    init: () => Promise<void>;
    setPixelRatio: (value: number) => void;
    setSize: (width: number, height: number, updateStyle?: boolean) => void;
    render: (scene: unknown, camera: unknown) => void;
    getContext?: () => WebGL2RenderingContext;
    backend?: { isWebGLBackend?: boolean };
  } | null = null;
  private readonly scene = new THREE.Scene();
  private readonly camera = new THREE.PerspectiveCamera(45, 1, 0.1, 100000);
  private readonly root = new THREE.Group();
  private readonly highlight: AnyObject;
  private drawn: Drawn[] = [];
  private walls: ShadowWalls | null = null;
  private paths: PickPath[] = [];
  private orbit: Orbit = { targetX: 0, targetY: 0, targetZ: 0, theta: -Math.PI / 4, phi: Math.PI / 3, radius: 10 };
  private home: Orbit = this.orbit;
  private width = 1;
  private height = 1;
  private dpr = 1;
  private dragging = false;
  private dragX = 0;
  private dragY = 0;
  private dragButton = 0;
  private dragShift = false;
  private showDrilled = true;
  /** Null until the page sends a filter: the concern wells, which is what the view opens on. */
  private shown: Set<string> | null = null;
  private preferTubes = false;
  private dirty = true;
  private paused = false;
  private lost = false;
  private cached: WellScene | null = null;
  private loopStarted = false;
  private readonly sphere: { dispose?: () => void };

  constructor(private readonly canvas: OffscreenCanvas) {
    this.scene.background = new THREE.Color(BACKGROUND);
    this.scene.add(this.root);
    this.camera.up.set(0, 0, 1);
    this.sphere = new THREE.SphereGeometry(1, 14, 10);
    this.highlight = new THREE.Mesh(this.sphere, new THREE.MeshBasicMaterial({ color: "#ffffff" }));
    this.highlight.visible = false;
    this.scene.add(this.highlight);
  }

  async init(forceWebGL: boolean, width: number, height: number, dpr: number): Promise<WellsBackend> {
    this.width = Math.max(1, width);
    this.height = Math.max(1, height);
    this.dpr = Math.min(Math.max(dpr, 1), 2);
    const renderer = new WebGPURenderer({
      canvas: this.canvas,
      antialias: false,
      alpha: false,
      powerPreference: "high-performance",
      forceWebGL,
    });
    await renderer.init();
    renderer.setPixelRatio(this.dpr);
    renderer.setSize(this.width, this.height, false);
    this.renderer = renderer;
    this.listenForContextLoss();
    this.startLoop();
    return this.backend();
  }

  resize(width: number, height: number, dpr: number): void {
    this.width = Math.max(1, width);
    this.height = Math.max(1, height);
    this.dpr = Math.min(Math.max(dpr, 1), 2);
    this.renderer?.setPixelRatio(this.dpr);
    this.renderer?.setSize(this.width, this.height, false);
    this.camera.aspect = this.width / this.height;
    this.camera.updateProjectionMatrix();
    this.walls?.fitLabels(this.camera.position, this.camera.fov, this.height);
    this.dirty = true;
  }

  /** A wide line across a white field, then the pixels a 1px line would miss. */
  probe(): ProbeResult {
    this.paused = true;
    this.clearRoot();
    const background = this.scene.background;
    this.scene.background = new THREE.Color("#ffffff");
    const camera = new THREE.OrthographicCamera(-1, 1, 1, -1, 0.1, 10);
    camera.position.set(0, 0, 5);
    camera.lookAt(0, 0, 0);
    const positions = new Float32Array([-0.9, 0, 0, 0.9, 0, 0]);
    let wide: ProbeResult["wide"] = "line";
    let object = this.makeStroke(positions, "#ff0000", 8, 1);
    this.scene.add(object);
    this.paint(camera);
    let center = this.readPixel(this.width / 2, this.height / 2);
    let shoulder = this.readPixel(this.width / 2, this.height / 2 - 3);
    if (!isRed(center) || !isRed(shoulder)) {
      wide = "tube";
      this.preferTubes = true;
      this.scene.remove(object);
      object = this.screenBar();
      this.scene.add(object);
      this.paint(camera);
      center = this.readPixel(this.width / 2, this.height / 2);
      shoulder = this.readPixel(this.width / 2, this.height / 2 - 3);
    }
    this.scene.remove(object);
    this.scene.background = background;
    this.paused = false;
    this.dirty = true;
    return { type: "probe", backend: this.backend(), wide, center, shoulder };
  }

  /** One frame of the current scene, for the blank-canvas check. */
  sample(): { center: [number, number, number]; corner: [number, number, number] } {
    this.paused = true;
    this.paint(this.camera);
    const center = this.readPixel(this.width / 2, this.height / 2);
    const corner = this.readPixel(4, 4);
    this.paused = false;
    this.dirty = true;
    return { center, corner };
  }

  setScene(scene: WellScene | null): void {
    this.cached = scene;
    this.rebuild();
  }

  setOptions(drilled: boolean, shown: readonly string[]): void {
    this.showDrilled = drilled;
    this.shown = new Set(shown);
    this.applyVisibility();
    this.dirty = true;
  }

  pointer(phase: "down" | "move" | "up" | "wheel", x: number, y: number, button: number, shift: boolean, deltaY: number): void {
    if (phase === "down") {
      this.dragging = true;
      this.dragX = x;
      this.dragY = y;
      this.dragButton = button;
      this.dragShift = shift;
      this.highlight.visible = false;
      return;
    }
    if (phase === "up") {
      this.dragging = false;
      return;
    }
    if (phase === "wheel") {
      this.orbit = wheelOrbit(this.orbit, deltaY);
      this.applyCamera();
      this.dirty = true;
      return;
    }
    if (!this.dragging) return;
    const dx = x - this.dragX;
    const dy = y - this.dragY;
    this.dragX = x;
    this.dragY = y;
    this.orbit = dragOrbit(this.orbit, dx, dy, this.dragButton === 2 || this.dragShift || shift);
    this.applyCamera();
    this.dirty = true;
  }

  hover(x: number, y: number) {
    if (this.dragging || this.paths.length === 0) return null;
    const hit = pickHover(this.visiblePaths(), { x, y }, (px, py, pz) => this.project(px, py, pz), 14);
    if (!hit) {
      this.highlight.visible = false;
      this.dirty = true;
      return null;
    }
    const path = this.visiblePaths().find((item) => item.kind === hit.kind && item.name === hit.name);
    const at = path?.points.find((point) => point.md === hit.refMd) ?? path?.points[0];
    if (at) {
      const radius = Math.max(this.orbit.radius * 0.012, 0.4);
      this.highlight.position.set(at.x, at.y, at.z);
      this.highlight.scale.setScalar(radius);
      this.highlight.visible = true;
      this.dirty = true;
    }
    return hit;
  }

  reset(): void {
    this.orbit = { ...this.home };
    this.applyCamera();
    this.dirty = true;
  }

  /** Steady frames after the first, which pays for shader compile. */
  bench(frames: number): { maxMs: number; frames: number } {
    this.paused = true;
    let maxMs = 0;
    const count = Math.max(2, frames);
    for (let i = 0; i < count; i += 1) {
      this.orbit = dragOrbit(this.orbit, 6, 2, false);
      this.applyCamera();
      const start = performance.now();
      this.renderer?.render(this.scene, this.camera);
      const ms = performance.now() - start;
      if (i > 0 && ms > maxMs) maxMs = ms;
    }
    this.orbit = { ...this.home };
    this.applyCamera();
    this.paused = false;
    this.dirty = true;
    return { maxMs, frames: count - 1 };
  }

  private rebuild(): void {
    this.clearRoot();
    this.paths = [];
    this.drawn = [];
    this.walls = null;
    const scene = this.cached;
    if (!scene) {
      this.dirty = true;
      return;
    }
    const bounds = emptyBounds();
    includeXyz(bounds, scene.plan.xyz);
    if (scene.drilled) includeXyz(bounds, scene.drilled.xyz);
    for (const offset of scene.offsets) if (offset.concern) includeXyz(bounds, offset.xyz);
    if (!Number.isFinite(bounds.minX)) includeXyz(bounds, scene.plan.xyz);
    const box = shadowBoxOf(scene);
    if (box) {
      this.walls = new ShadowWalls(box, scene.plan.xyz, scene.drilled?.xyz ?? null, scene.formations, (xyz, color, width, opacity) => this.makeStroke(xyz, color, width, opacity));
      this.root.add(this.walls.group);
      includeXyz(bounds, [box.minX, box.minY, box.minZ, box.maxX, box.maxY, box.maxZ]);
    }
    this.home = fitOrbit(bounds);
    this.orbit = { ...this.home };
    this.applyCamera();
    this.addPath(scene.plan.xyz, "plan", "Planned Path", PLAN_COLOR, PATH_WIDTH, 1, true, scene.plan);
    if (scene.drilled) this.addPath(scene.drilled.xyz, "drilled", "Drilled Path", DRILLED_COLOR, PATH_WIDTH, 1, true, scene.drilled);
    for (const offset of scene.offsets) this.addOffset(offset);
    for (const target of scene.targets) this.addTarget(target);
    this.applyVisibility();
    this.dirty = true;
  }

  private addPath(xyz: Float32Array, kind: Drawn["kind"], name: string, color: string, width: number, opacity: number, concern: boolean, packed: OffsetPath | WellScene["plan"], breaks: number[] = []): void {
    for (const run of runsOf(xyz, breaks)) {
      if (run.length < 6) continue;
      const object = this.makeStroke(run, color, width, opacity);
      object.renderOrder = kind === "plan" ? 1 : 2;
      this.root.add(object);
      this.drawn.push({ object, kind, concern, name });
    }
    this.paths.push({ kind, name, points: unpackPath(packed), breaks });
  }

  private addOffset(offset: OffsetPath): void {
    const width = offset.concern ? 5 : 2;
    const opacity = offset.concern ? 1 : 0.4;
    this.addPath(offset.xyz, "offset", offset.name, offset.color, width, opacity, offset.concern, offset, offset.breaks);
  }

  /** Two rings on the plan, square to the wellbore, at a named target. */
  private addTarget(target: { name: string; md: number; tvd: number; x: number; y: number; z: number; outer: Float32Array; inner: Float32Array }): void {
    for (const ring of [target.outer, target.inner]) {
      const object = this.makeStroke(ring, TARGET_COLOR, 2, 1);
      this.root.add(object);
      this.drawn.push({ object, kind: "plan", concern: true, name: target.name });
    }
    this.paths.push({
      kind: "target",
      name: target.name,
      points: [{ x: target.x, y: target.y, z: target.z, md: target.md, tvd: target.tvd, c2c: null, sf: null }],
    });
  }

  private applyVisibility(): void {
    for (const item of this.drawn) {
      if (item.kind === "drilled") item.object.visible = this.showDrilled;
      else if (item.kind === "offset") item.object.visible = this.shown === null ? item.concern : this.shown.has(item.name);
      else item.object.visible = true;
    }
    this.walls?.showDrilled(this.showDrilled);
  }

  private visiblePaths(): PickPath[] {
    return this.paths.filter((path) => {
      if (path.kind === "drilled") return this.showDrilled;
      if (path.kind === "offset") {
        if (this.shown) return this.shown.has(path.name);
        const offset = this.cached?.offsets.find((item) => item.name === path.name);
        return offset?.concern ?? false;
      }
      return true;
    });
  }

  private makeStroke(xyz: Float32Array, color: string, width: number, opacity: number): AnyObject {
    if (this.preferTubes) return this.tube(xyz, color, opacity);
    try {
      const geometry = new LineGeometry();
      geometry.setPositions(xyz);
      // Always in the transparent pass, with the walls, so render order puts every line after them
      // and a wall never washes out a well behind it.
      const material = new Line2NodeMaterial({
        color,
        linewidth: width,
        transparent: true,
        opacity,
        worldUnits: false,
      });
      material.linewidth = width;
      return new Line2(geometry, material);
    } catch {
      this.preferTubes = true;
      return this.tube(xyz, color, opacity);
    }
  }

  private tube(xyz: Float32Array, color: string, opacity: number): AnyObject {
    const count = Math.max(0, xyz.length / 3 - 1);
    const geometry = new THREE.CylinderGeometry(0.45, 0.45, 1, 5);
    const material = new THREE.MeshBasicMaterial({ color, transparent: true, opacity });
    const mesh = new THREE.InstancedMesh(geometry, material, Math.max(1, count));
    const dummy = new THREE.Object3D();
    const up = new THREE.Vector3(0, 1, 0);
    const dir = new THREE.Vector3();
    for (let i = 0; i < count; i += 1) {
      const ax = xyz[i * 3]!;
      const ay = xyz[i * 3 + 1]!;
      const az = xyz[i * 3 + 2]!;
      const bx = xyz[i * 3 + 3]!;
      const by = xyz[i * 3 + 4]!;
      const bz = xyz[i * 3 + 5]!;
      dir.set(bx - ax, by - ay, bz - az);
      const length = dir.length() || 0.001;
      dir.divideScalar(length);
      dummy.position.set((ax + bx) / 2, (ay + by) / 2, (az + bz) / 2);
      dummy.quaternion.setFromUnitVectors(up, dir);
      dummy.scale.set(1, length, 1);
      dummy.updateMatrix();
      mesh.setMatrixAt(i, dummy.matrix);
    }
    mesh.instanceMatrix.needsUpdate = true;
    return mesh;
  }

  private screenBar(): AnyObject {
    const height = 8 / Math.max(this.height, 1) * 2;
    return new THREE.Mesh(new THREE.BoxGeometry(1.8, height, 0.01), new THREE.MeshBasicMaterial({ color: "#ff0000" }));
  }

  private project(x: number, y: number, z: number): { x: number; y: number; behind: boolean } | null {
    const vector = new THREE.Vector3(x, y, z).project(this.camera);
    if (!Number.isFinite(vector.x) || !Number.isFinite(vector.y)) return null;
    return {
      x: (vector.x * 0.5 + 0.5) * this.width,
      y: (-vector.y * 0.5 + 0.5) * this.height,
      behind: vector.z < -1 || vector.z > 1,
    };
  }

  private applyCamera(): void {
    const eye = cameraFromOrbit(this.orbit);
    this.camera.aspect = this.width / Math.max(1, this.height);
    this.camera.position.set(eye.x, eye.y, eye.z);
    this.camera.up.set(0, 0, 1);
    this.camera.lookAt(this.orbit.targetX, this.orbit.targetY, this.orbit.targetZ);
    this.walls?.follow(eye);
    this.walls?.fitLabels(eye, this.camera.fov, this.height);
    this.camera.near = Math.max(0.05, this.orbit.radius / 500);
    this.camera.far = Math.max(100, this.orbit.radius * 40);
    this.camera.updateProjectionMatrix();
  }

  private render(camera: unknown = this.camera): void {
    if (this.lost) return;
    this.renderer?.render(this.scene, camera);
  }

  /** The first draw pays for shader compile, so a proof reads the second. */
  private paint(camera: unknown): void {
    this.render(camera);
    this.render(camera);
  }

  private readPixel(x: number, y: number): [number, number, number] {
    const gl = this.renderer?.getContext?.() as WebGL2RenderingContext | undefined;
    if (!gl || typeof gl.readPixels !== "function") return [0, 0, 0];
    const bufferW = Math.max(1, this.canvas.width);
    const bufferH = Math.max(1, this.canvas.height);
    gl.bindFramebuffer(gl.FRAMEBUFFER, null);
    const buffer = new Uint8Array(4);
    const px = Math.max(0, Math.min(bufferW - 1, Math.floor(x * bufferW / this.width)));
    const py = Math.max(0, Math.min(bufferH - 1, Math.floor(bufferH - (y * bufferH / this.height) - 1)));
    gl.readPixels(px, py, 1, 1, gl.RGBA, gl.UNSIGNED_BYTE, buffer);
    return [buffer[0]!, buffer[1]!, buffer[2]!];
  }

  private backend(): WellsBackend {
    return this.renderer?.backend?.isWebGLBackend ? "webgl2" : "webgpu";
  }

  private listenForContextLoss(): void {
    this.canvas.addEventListener("webglcontextlost", (event) => {
      event.preventDefault();
      this.lost = true;
    });
    this.canvas.addEventListener("webglcontextrestored", () => {
      this.lost = false;
      this.rebuild();
    });
  }

  private startLoop(): void {
    if (this.loopStarted) return;
    this.loopStarted = true;
    const tick = () => {
      requestAnimationFrame(tick);
      if (this.paused || !this.dirty || this.lost) return;
      this.dirty = false;
      this.render();
    };
    requestAnimationFrame(tick);
  }

  private clearRoot(): void {
    for (const child of [...this.root.children]) {
      this.root.remove(child);
      disposeNode(child);
    }
  }
}

function disposeNode(node: AnyObject): void {
  const mesh = node;
  if (mesh.geometry && mesh.geometry.type !== "SphereGeometry") mesh.geometry.dispose?.();
  const materials = Array.isArray(mesh.material) ? mesh.material : [mesh.material];
  for (const material of materials) {
    const map = (material as { map?: { dispose?: () => void; userData?: { disposed?: boolean } } } | undefined)?.map;
    if (map && !map.userData?.disposed) {
      map.userData = { ...(map.userData ?? {}), disposed: true };
      map.dispose?.();
    }
    material?.dispose?.();
  }
  for (const child of [...node.children]) disposeNode(child);
}

function isRed(pixel: [number, number, number]): boolean {
  return pixel[0] > 180 && pixel[1] < 80 && pixel[2] < 80;
}
