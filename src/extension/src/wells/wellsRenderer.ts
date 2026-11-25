import * as THREE from "three";
import { Line2NodeMaterial, WebGPURenderer } from "three/webgpu";
import { Line2 } from "three/addons/lines/webgpu/Line2.js";
import { LineGeometry } from "three/addons/lines/LineGeometry.js";
import { dragOrbit, emptyBounds, fitOrbit, includeXyz, leashOrbit, wheelOrbit, cameraFromOrbit, type Orbit } from "../../../packages/wells-3d/src/orbit";
import { shadowBoxOf } from "../../../packages/wells-3d/src/shadowBox";
import { pickHover, unpackPath, type PickPath } from "../../../packages/wells-3d/src/wellPick";
import { ringLabelPoint, ringMetresLabel } from "../../../packages/wells-3d/src/ringLabel";
import type { OffsetPath, WellScene } from "../../../packages/wells-3d/src/wellScene";
import type { ProbeResult, WellsBackend } from "./wellsMessages";
import type { AnyObject } from "./wellsObject";
import { ShadowWalls } from "./wellsWalls";

const PLAN_COLOR = "#17324d";
const DRILLED_COLOR = "#0e7490";
const TARGET_COLOR = "#1b4d2e";
// Screen pixels. Thin station-to-station legs, the way a directional plot draws a plan and a survey.
// The active well stays heavier than the offsets so the pair being compared still reads first.
const PLAN_WIDTH = 4;
const DRILLED_WIDTH = 6;
const OFFSET_WIDTH = 3;
const OFFSET_DIM_WIDTH = 1.75;
const PLAN_STATION = 8;
const DRILLED_STATION = 9;
/** A leg longer than this is split, so the view cannot throw the whole leg away when one end leaves the frame. */
const MAX_LEG_M = 20;
const BACKGROUND = "#d5dee6";
/** The distance words stay this tall on screen, and hide when the ring is smaller than them. */
const RING_LABEL_PX = 15;
const RING_LABEL_MIN_PX = 22;

interface Drawn {
  object: AnyObject;
  kind: "plan" | "offset" | "drilled";
  concern: boolean;
  name: string;
}

interface RingLabel {
  mesh: AnyObject;
  ring: Float32Array;
  center: { x: number; y: number; z: number };
  pxWidth: number;
  pxHeight: number;
}

/** Bold metres, the ring's own green, with a light edge so it reads on the grid. */
function paintRingLabel(text: string): { canvas: OffscreenCanvas; pxWidth: number; pxHeight: number } | null {
  const size = 64;
  const pad = 12;
  const measure = new OffscreenCanvas(4, 4);
  const ruler = measure.getContext("2d");
  if (!ruler) return null;
  ruler.font = `bold ${size}px sans-serif`;
  const width = Math.max(4, Math.ceil(ruler.measureText(text).width + pad * 2));
  const height = Math.max(4, Math.ceil(size + pad * 2));
  const canvas = new OffscreenCanvas(width, height);
  const ctx = canvas.getContext("2d");
  if (!ctx) return null;
  ctx.font = `bold ${size}px sans-serif`;
  ctx.textBaseline = "middle";
  ctx.lineWidth = 10;
  ctx.lineJoin = "round";
  ctx.strokeStyle = "rgba(255,255,255,0.94)";
  ctx.fillStyle = TARGET_COLOR;
  ctx.strokeText(text, pad, height / 2);
  ctx.fillText(text, pad, height / 2);
  return { canvas, pxWidth: width, pxHeight: height };
}

/** Insert points so no leg is long enough to vanish when one end is outside the view. A repeated point is dropped. */
function densify(xyz: Float32Array): Float32Array {
  const count = Math.floor(xyz.length / 3);
  if (count < 2) return xyz;
  const out: number[] = [xyz[0]!, xyz[1]!, xyz[2]!];
  for (let i = 1; i < count; i += 1) {
    const ax = out[out.length - 3]!;
    const ay = out[out.length - 2]!;
    const az = out[out.length - 1]!;
    const bx = xyz[i * 3]!;
    const by = xyz[i * 3 + 1]!;
    const bz = xyz[i * 3 + 2]!;
    const dx = bx - ax;
    const dy = by - ay;
    const dz = bz - az;
    const len = Math.hypot(dx, dy, dz);
    if (len < 0.05) continue;
    const pieces = Math.max(1, Math.ceil(len / MAX_LEG_M));
    for (let step = 1; step <= pieces; step += 1) {
      const t = step / pieces;
      out.push(ax + dx * t, ay + dy * t, az + dz * t);
    }
  }
  return out.length < 6 ? xyz : Float32Array.from(out);
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
  private ringLabels: RingLabel[] = [];
  private paths: PickPath[] = [];
  private orbit: Orbit = { targetX: 0, targetY: 0, targetZ: 0, theta: -Math.PI / 4, phi: Math.PI / 3, radius: 10 };
  private home: Orbit = this.orbit;
  /** True after the first scene. Later rebuilds keep this angle; Reset View still uses home. */
  private framed = false;
  /** Plan and drilled polylines. The look-at stays on these, so the active well stays in frame. */
  private focusPaths: ArrayLike<number>[] = [];
  /** Longest side of everything drawn, so a close zoom does not clip the rest of the field. */
  private extent = 100;
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
    const windowsWebGpu = !forceWebGL && typeof navigator !== "undefined" && /Windows/i.test(navigator.userAgent);
    const renderer = new WebGPURenderer({
      canvas: this.canvas,
      antialias: false,
      alpha: false,
      // Chromium ignores powerPreference on Windows WebGPU and logs crbug.com/369219127 as an extension error.
      ...(windowsWebGpu ? {} : { powerPreference: "high-performance" }),
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
    this.fitRingLabels();
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
      this.orbit = this.holdOnWell(wheelOrbit(this.orbit, deltaY));
      this.applyCamera();
      this.dirty = true;
      return;
    }
    if (!this.dragging) return;
    const dx = x - this.dragX;
    const dy = y - this.dragY;
    this.dragX = x;
    this.dragY = y;
    const pan = this.dragButton === 2 || this.dragShift || shift;
    this.orbit = dragOrbit(this.orbit, dx, dy, pan);
    // Left drag only turns. Right drag slides, then stops once the active well would leave the frame.
    if (pan) this.orbit = this.holdOnWell(this.orbit);
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
    const kept = this.framed ? { ...this.orbit } : null;
    this.clearRoot();
    this.paths = [];
    this.drawn = [];
    this.ringLabels = [];
    this.walls = null;
    this.focusPaths = [];
    const scene = this.cached;
    if (!scene) {
      this.framed = false;
      this.dirty = true;
      return;
    }
    const bounds = emptyBounds();
    includeXyz(bounds, scene.plan.xyz);
    if (scene.drilled) includeXyz(bounds, scene.drilled.xyz);
    const field = emptyBounds();
    includeXyz(field, scene.plan.xyz);
    if (scene.drilled) includeXyz(field, scene.drilled.xyz);
    for (const offset of scene.offsets) includeXyz(field, offset.xyz);
    this.extent = Math.max(field.maxX - field.minX, field.maxY - field.minY, field.maxZ - field.minZ, 100);
    const box = shadowBoxOf(scene);
    if (box) {
      this.walls = new ShadowWalls(box, scene.plan.xyz, scene.drilled?.xyz ?? null, scene.formations, (xyz, color, width, opacity) => this.makeStroke(xyz, color, width, opacity));
      this.root.add(this.walls.group);
    }
    this.focusPaths = scene.drilled ? [scene.plan.xyz, scene.drilled.xyz] : [scene.plan.xyz];
    // Frame the active well. The look-at is its centre, already inside the pan leash, so Reset View does not jump.
    this.home = this.holdOnWell(fitOrbit(bounds));
    // A survey poll rebuilds the picture while the user is on Sheets. Keep the angle they set.
    this.orbit = kept ? this.holdOnWell(kept) : { ...this.home };
    this.framed = true;
    this.applyCamera();
    this.addPath(scene.plan.xyz, "plan", "Planned Path", PLAN_COLOR, PLAN_WIDTH, 1, true, scene.plan);
    this.addStations(scene.plan.xyz, "plan", "Planned Path", PLAN_COLOR, PLAN_STATION, true);
    if (scene.drilled) {
      this.addPath(scene.drilled.xyz, "drilled", "Drilled Path", DRILLED_COLOR, DRILLED_WIDTH, 1, true, scene.drilled);
      this.addStations(scene.drilled.xyz, "drilled", "Drilled Path", DRILLED_COLOR, DRILLED_STATION, true);
    }
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

  /** How far the look-at may sit from the active well at this zoom. Sliding along the well is free. */
  private holdOnWell(orbit: Orbit): Orbit {
    return leashOrbit(orbit, this.focusPaths, Math.max(40, orbit.radius * 0.35));
  }

  /** A bead on every plan and survey station, so each leg reads as a segment. */
  private addStations(xyz: Float32Array, kind: Drawn["kind"], name: string, color: string, size: number, concern: boolean): void {
    if (xyz.length < 3) return;
    const geometry = new THREE.BufferGeometry();
    geometry.setAttribute("position", new THREE.BufferAttribute(xyz, 3));
    geometry.computeBoundingSphere();
    const material = new THREE.PointsMaterial({
      color,
      size,
      sizeAttenuation: false,
      transparent: true,
      opacity: 1,
      depthWrite: false,
    });
    const points = new THREE.Points(geometry, material);
    points.frustumCulled = false;
    points.renderOrder = 4;
    this.root.add(points);
    this.drawn.push({ object: points, kind, concern, name });
  }

  private addOffset(offset: OffsetPath): void {
    const width = offset.concern ? OFFSET_WIDTH : OFFSET_DIM_WIDTH;
    const opacity = offset.concern ? 1 : 0.4;
    this.addPath(offset.xyz, "offset", offset.name, offset.color, width, opacity, offset.concern, offset, offset.breaks);
  }

  /** Two rings on the plan, square to the wellbore, at a named target. Each ring carries its radius. */
  private addTarget(target: { name: string; md: number; tvd: number; x: number; y: number; z: number; outer: Float32Array; inner: Float32Array }): void {
    const center = { x: target.x, y: target.y, z: target.z };
    for (const ring of [target.outer, target.inner]) {
      const object = this.makeStroke(ring, TARGET_COLOR, 2, 1);
      this.root.add(object);
      this.drawn.push({ object, kind: "plan", concern: true, name: target.name });
      const metres = Math.round(Math.hypot((ring[0] ?? center.x) - center.x, (ring[1] ?? center.y) - center.y, (ring[2] ?? center.z) - center.z));
      const painted = metres > 0.05 ? paintRingLabel(ringMetresLabel(metres)) : null;
      if (!painted) continue;
      const texture = new THREE.CanvasTexture(painted.canvas);
      texture.colorSpace = THREE.SRGBColorSpace;
      texture.generateMipmaps = false;
      texture.minFilter = THREE.LinearFilter;
      texture.magFilter = THREE.LinearFilter;
      texture.needsUpdate = true;
      const mesh = new THREE.Sprite(new THREE.SpriteMaterial({
        map: texture,
        transparent: true,
        depthWrite: false,
        depthTest: false,
      }));
      mesh.frustumCulled = false;
      mesh.renderOrder = 6;
      mesh.visible = false;
      this.root.add(mesh);
      this.ringLabels.push({ mesh, ring, center, pxWidth: painted.pxWidth, pxHeight: painted.pxHeight });
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
    const legs = densify(xyz);
    if (this.preferTubes) return this.tube(legs, color, opacity);
    try {
      const geometry = new LineGeometry();
      geometry.setPositions(legs);
      // Always in the transparent pass, with the walls, so render order puts every line after them
      // and a wall never washes out a well behind it. Depth is not written, so one well cannot
      // erase the section of another where the two cross on screen.
      const material = new Line2NodeMaterial({
        color,
        linewidth: width,
        transparent: true,
        opacity,
        depthWrite: false,
        worldUnits: false,
      });
      material.linewidth = width;
      const line = new Line2(geometry, material);
      line.frustumCulled = false;
      return line;
    } catch {
      this.preferTubes = true;
      return this.tube(legs, color, opacity);
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
    mesh.frustumCulled = false;
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
    const reach = Math.max(this.extent, this.orbit.radius);
    this.camera.near = Math.max(0.25, this.orbit.radius / 2000);
    this.camera.far = Math.max(reach * 4, this.orbit.radius * 80);
    this.camera.updateProjectionMatrix();
    this.fitRingLabels();
  }

  /**
   * Each ring's radius, in metres, at the bottom of the ring the way a target plot writes 7" / 17" / 27".
   * The words sit just inside the stroke and stay a fixed size on screen.
   */
  private fitRingLabels(): void {
    if (this.ringLabels.length === 0) return;
    this.camera.updateMatrixWorld(true);
    const half = Math.tan((this.camera.fov * Math.PI) / 180 / 2);
    const viewHeight = Math.max(1, this.height);
    for (const label of this.ringLabels) {
      const rim = this.bottomOf(label.ring, label.center);
      if (!rim) {
        label.mesh.visible = false;
        continue;
      }
      const distance = Math.max(1, Math.hypot(this.camera.position.x - rim.x, this.camera.position.y - rim.y, this.camera.position.z - rim.z));
      const worldPerPx = (2 * half * distance) / viewHeight;
      const worldH = RING_LABEL_PX * worldPerPx;
      const at = ringLabelPoint(label.center, rim, worldH * 0.65);
      label.mesh.position.set(at.x, at.y, at.z);
      label.mesh.scale.set(worldH * (label.pxWidth / label.pxHeight), worldH, 1);
      label.mesh.visible = true;
    }
  }

  /** The point on the ring that lands lowest on screen, when the ring is large enough to read. */
  private bottomOf(ring: Float32Array, center: { x: number; y: number; z: number }): { x: number; y: number; z: number } | null {
    const origin = this.project(center.x, center.y, center.z);
    if (!origin || origin.behind) return null;
    let bestY = -Infinity;
    let best: { x: number; y: number; z: number } | null = null;
    let bestScreen = 0;
    const count = Math.floor(ring.length / 3);
    const last = count > 1 ? count - 1 : count;
    for (let i = 0; i < last; i += 1) {
      const x = ring[i * 3]!;
      const y = ring[i * 3 + 1]!;
      const z = ring[i * 3 + 2]!;
      const screen = this.project(x, y, z);
      if (!screen || screen.behind) continue;
      const reach = Math.hypot(screen.x - origin.x, screen.y - origin.y);
      if (reach > bestScreen) bestScreen = reach;
      if (screen.y > bestY) {
        bestY = screen.y;
        best = { x, y, z };
      }
    }
    if (!best || bestScreen < RING_LABEL_MIN_PX) return null;
    return best;
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
  // Sprite geometry is one shared buffer. Disposing it blanks every later label.
  const shared = mesh.geometry?.type === "SphereGeometry" || (mesh as { isSprite?: boolean }).isSprite === true;
  if (mesh.geometry && !shared) mesh.geometry.dispose?.();
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
