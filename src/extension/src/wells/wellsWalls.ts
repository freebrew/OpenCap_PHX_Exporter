import * as THREE from "three";
import { formationMarks, formationTvdLabel } from "../../../packages/wells-3d/src/formationMarks";
import { flattenOnto, gridValues, wallsFor, type ShadowBox } from "../../../packages/wells-3d/src/shadowBox";
import type { WellFormation } from "../../../packages/wells-3d/src/wellScene";
import type { AnyObject, StrokeMaker } from "./wellsObject";

// Colours of the email picture's shadow box (corridorDraw.ts): white walls, a grey floor, the grid
// and the shadow in greys. The walls are see-through, so an offset well behind one is still seen.
const WALL_FILL = "#ffffff";
const FLOOR_FILL = "#e8ebee";
const WALL_EDGE = "#5c656c";
const GRID = "#b4bcc2";
const SHADOW = "#7a848c";
const SHADOW_WIDTH = 2;
const FILL_OPACITY = 0.6;
// The email picture's formation ink (corridorDraw.ts drawFormationsOnShaft).
const FORMATION = "#6b4f2a";
const LABEL_SCREEN_PX = 18;

type Corner = [number, number, number];

/** Behind everything the data draws. The shadow comes last of the three. */
const ORDER_FILL = -30;
const ORDER_LINES = -20;
const ORDER_FORMATION = -15;
const ORDER_LABEL = -5;
const ORDER_SHADOW = -10;

function quad(corners: [Corner, Corner, Corner, Corner], color: string): AnyObject {
  const geometry = new THREE.BufferGeometry();
  geometry.setAttribute("position", new THREE.Float32BufferAttribute(corners.flat(), 3));
  geometry.setIndex([0, 1, 2, 0, 2, 3]);
  const mesh: AnyObject = new THREE.Mesh(geometry, new THREE.MeshBasicMaterial({
    color,
    transparent: true,
    opacity: FILL_OPACITY,
    depthWrite: false,
    side: THREE.DoubleSide,
  }));
  mesh.renderOrder = ORDER_FILL;
  return mesh;
}

function segments(positions: number[], color: string): AnyObject {
  const geometry = new THREE.BufferGeometry();
  geometry.setAttribute("position", new THREE.Float32BufferAttribute(positions, 3));
  const lines: AnyObject = new THREE.LineSegments(geometry, new THREE.LineBasicMaterial({ color, transparent: true, opacity: 0.95, depthWrite: false }));
  lines.renderOrder = ORDER_LINES;
  return lines;
}

function frame(corners: [Corner, Corner, Corner, Corner]): number[] {
  const out: number[] = [];
  for (let i = 0; i < 4; i += 1) out.push(...corners[i]!, ...corners[(i + 1) % 4]!);
  return out;
}

type Axis = "x" | "y" | "z";

/**
 * One face of the box with its grid. `fixed` is the axis the face is square to, `at` where it sits,
 * and `across` the two axes the face spreads over.
 */
function face(box: ShadowBox, fixed: Axis, at: number): { fill: AnyObject; grid: AnyObject; edge: AnyObject } {
  const range: Record<Axis, [number, number]> = { x: [box.minX, box.maxX], y: [box.minY, box.maxY], z: [box.minZ, box.maxZ] };
  const across = (["x", "y", "z"] as const).filter((axis) => axis !== fixed);
  const [a, b] = across as [Axis, Axis];
  const place = (u: number, v: number): Corner => {
    const point: Record<Axis, number> = { x: 0, y: 0, z: 0 };
    point[fixed] = at;
    point[a] = u;
    point[b] = v;
    return [point.x, point.y, point.z];
  };
  const corners: [Corner, Corner, Corner, Corner] = [
    place(range[a][0], range[b][0]),
    place(range[a][1], range[b][0]),
    place(range[a][1], range[b][1]),
    place(range[a][0], range[b][1]),
  ];
  const lines: number[] = [];
  for (const u of gridValues(range[a][0], range[a][1], box.step)) lines.push(...place(u, range[b][0]), ...place(u, range[b][1]));
  for (const v of gridValues(range[b][0], range[b][1], box.step)) lines.push(...place(range[a][0], v), ...place(range[a][1], v));
  return {
    fill: quad(corners, fixed === "z" ? FLOOR_FILL : WALL_FILL),
    grid: segments(lines, GRID),
    edge: segments(frame(corners), WALL_EDGE),
  };
}

type WallAxis = "x" | "y";
type WallSideName = "min" | "max";

interface VerticalWall {
  node: AnyObject;
  fixed: WallAxis;
  at: number;
  side: WallSideName;
  alongMin: number;
  alongMax: number;
  /** World direction of left-to-right as this wall is read. It faces the camera that can see the wall. */
  right: [number, number, number];
  /** Toward that camera, so the label sits just off the wall. */
  normal: [number, number, number];
}

interface WallLabel {
  mesh: AnyObject;
  z: number;
  along: number;
  pxWidth: number;
  pxHeight: number;
  wall: VerticalWall;
}

/** Which way the writing runs, and which way is toward the camera, for the one side that wall is seen from. */
function wallFacing(fixed: WallAxis, side: WallSideName): { right: [number, number, number]; normal: [number, number, number] } {
  switch (fixed) {
    case "y":
      return side === "min"
        ? { right: [-1, 0, 0], normal: [0, 1, 0] }
        : { right: [1, 0, 0], normal: [0, -1, 0] };
    case "x":
      return side === "min"
        ? { right: [0, 1, 0], normal: [1, 0, 0] }
        : { right: [0, -1, 0], normal: [-1, 0, 0] };
    default: {
      const never: never = fixed;
      return never;
    }
  }
}

function paintLabel(name: string, tvd: number): { canvas: OffscreenCanvas; pxWidth: number; pxHeight: number } | null {
  const font = "sans-serif";
  const size = 64;
  const pad = 14;
  const gap = 36;
  const tvdText = formationTvdLabel(tvd);
  const measure = new OffscreenCanvas(4, 4);
  const ruler = measure.getContext("2d");
  if (!ruler) return null;
  ruler.font = `bold ${size}px ${font}`;
  const nameW = ruler.measureText(name).width;
  ruler.font = `${size}px ${font}`;
  const tvdW = ruler.measureText(tvdText).width;
  const width = Math.max(4, Math.ceil(pad * 2 + nameW + gap + tvdW));
  const height = Math.max(4, Math.ceil(size + pad * 2));
  const canvas = new OffscreenCanvas(width, height);
  const ctx = canvas.getContext("2d");
  if (!ctx) return null;
  ctx.textBaseline = "middle";
  ctx.lineWidth = 8;
  ctx.lineJoin = "round";
  ctx.strokeStyle = "rgba(255,255,255,0.94)";
  ctx.fillStyle = FORMATION;
  const y = height / 2;
  ctx.font = `bold ${size}px ${font}`;
  ctx.strokeText(name, pad, y);
  ctx.fillText(name, pad, y);
  ctx.font = `${size}px ${font}`;
  const tvdX = pad + nameW + gap;
  ctx.strokeText(tvdText, tvdX, y);
  ctx.fillText(tvdText, tvdX, y);
  return { canvas, pxWidth: width, pxHeight: height };
}

function labelTexture(canvas: OffscreenCanvas): AnyObject {
  const texture = new THREE.CanvasTexture(canvas);
  texture.colorSpace = THREE.SRGBColorSpace;
  texture.generateMipmaps = false;
  texture.minFilter = THREE.LinearFilter;
  texture.magFilter = THREE.LinearFilter;
  texture.needsUpdate = true;
  return texture;
}

/** The active well's shadow box. Every wall is built once; `follow` shows the ones behind the well. */
export class ShadowWalls {
  readonly group: AnyObject = new THREE.Group();
  private readonly floor: AnyObject = new THREE.Group();
  private readonly xMin: AnyObject = new THREE.Group();
  private readonly xMax: AnyObject = new THREE.Group();
  private readonly yMin: AnyObject = new THREE.Group();
  private readonly yMax: AnyObject = new THREE.Group();
  private readonly drilledShadows: AnyObject[] = [];
  private readonly labels: WallLabel[] = [];
  private readonly inset: number;

  constructor(private readonly box: ShadowBox, plan: Float32Array, drilled: Float32Array | null, formations: readonly WellFormation[], stroke: StrokeMaker) {
    const faces: Array<{ node: AnyObject; fixed: Axis; at: number }> = [
      { node: this.floor, fixed: "z", at: box.minZ },
      { node: this.xMin, fixed: "x", at: box.minX },
      { node: this.xMax, fixed: "x", at: box.maxX },
      { node: this.yMin, fixed: "y", at: box.minY },
      { node: this.yMax, fixed: "y", at: box.maxY },
    ];
    for (const item of faces) {
      const drawn = face(box, item.fixed, item.at);
      item.node.add(drawn.fill);
      item.node.add(drawn.grid);
      item.node.add(drawn.edge);
      this.shadowOn(item.node, plan, item.fixed, item.at, stroke, false);
      if (drilled) this.shadowOn(item.node, drilled, item.fixed, item.at, stroke, true);
      this.group.add(item.node);
    }
    const span = Math.max(box.maxX - box.minX, box.maxY - box.minY, box.maxZ - box.minZ);
    this.inset = Math.max(0.5, span * 0.0015);
    this.addFormations(formations);
  }

  /** Dashed tops and a name with its TVD, on every vertical wall. The wall's own visibility hides them. */
  private addFormations(formations: readonly WellFormation[]): void {
    const marks = formationMarks(this.box.minZ, this.box.maxZ, formations);
    if (marks.length === 0) return;
    const walls: VerticalWall[] = [
      this.vertical(this.xMin, "x", "min"),
      this.vertical(this.xMax, "x", "max"),
      this.vertical(this.yMin, "y", "min"),
      this.vertical(this.yMax, "y", "max"),
    ];
    for (const wall of walls) {
      const dashes: number[] = [];
      for (const mark of marks) dashes.push(...this.dash(wall, mark.z));
      const line = segments(dashes, FORMATION);
      line.renderOrder = ORDER_FORMATION;
      wall.node.add(line);
      for (const mark of marks) {
        const painted = paintLabel(mark.name, mark.tvd);
        if (!painted) continue;
        const material = new THREE.MeshBasicMaterial({
          map: labelTexture(painted.canvas),
          transparent: true,
          depthWrite: false,
          polygonOffset: true,
          polygonOffsetFactor: -4,
          polygonOffsetUnits: -4,
        });
        const mesh = new THREE.Mesh(new THREE.PlaneGeometry(1, 1), material);
        const basis = new THREE.Matrix4().makeBasis(
          new THREE.Vector3(wall.right[0], wall.right[1], wall.right[2]),
          new THREE.Vector3(0, 0, 1),
          new THREE.Vector3(wall.normal[0], wall.normal[1], wall.normal[2]),
        );
        mesh.quaternion.setFromRotationMatrix(basis);
        mesh.renderOrder = ORDER_LABEL;
        wall.node.add(mesh);
        this.labels.push({ mesh, z: mark.z, along: mark.along, pxWidth: painted.pxWidth, pxHeight: painted.pxHeight, wall });
      }
    }
  }

  private vertical(node: AnyObject, fixed: WallAxis, side: WallSideName): VerticalWall {
    const facing = wallFacing(fixed, side);
    const alongMin = fixed === "x" ? this.box.minY : this.box.minX;
    const alongMax = fixed === "x" ? this.box.maxY : this.box.maxX;
    const at = fixed === "x" ? (side === "min" ? this.box.minX : this.box.maxX) : (side === "min" ? this.box.minY : this.box.maxY);
    return { node, fixed, at, side, alongMin, alongMax, right: facing.right, normal: facing.normal };
  }

  /** A dashed line across the wall at one TVD. The dash is long enough to read at the fitted view. */
  private dash(wall: VerticalWall, z: number): number[] {
    const span = wall.alongMax - wall.alongMin;
    const dash = Math.max(8, span / 28);
    const gap = dash * 0.65;
    const out: number[] = [];
    const place = (u: number): number[] => (wall.fixed === "x" ? [wall.at, u, z] : [u, wall.at, z]);
    for (let cursor = wall.alongMin, guard = 0; cursor < wall.alongMax && guard < 400; guard += 1) {
      const end = Math.min(wall.alongMax, cursor + dash);
      out.push(...place(cursor), ...place(end));
      cursor = end + gap;
    }
    return out;
  }

  private place(wall: VerticalWall, along: number, z: number, lift: boolean): [number, number, number] {
    const pointsPositive = wall.fixed === "x" ? wall.right[1] > 0 : wall.right[0] > 0;
    const left = pointsPositive ? wall.alongMin : wall.alongMax;
    const rightEnd = pointsPositive ? wall.alongMax : wall.alongMin;
    const u = left + (rightEnd - left) * along;
    const inset = lift ? this.inset : 0;
    if (wall.fixed === "x") return [wall.at + wall.normal[0] * inset, u, z];
    return [u, wall.at + wall.normal[1] * inset, z];
  }

  private shadowOn(node: AnyObject, xyz: Float32Array, fixed: Axis, at: number, stroke: StrokeMaker, isDrilled: boolean): void {
    if (xyz.length < 6) return;
    const line = stroke(flattenOnto(xyz, fixed, at), SHADOW, SHADOW_WIDTH, 0.95);
    line.renderOrder = ORDER_SHADOW;
    node.add(line);
    if (isDrilled) this.drilledShadows.push(line);
  }

  /** The drilled path's shadow follows the drilled path's switch. */
  showDrilled(show: boolean): void {
    for (const line of this.drilledShadows) line.visible = show;
  }

  /** Show the walls that are behind the well from this camera. */
  follow(eye: { x: number; y: number; z: number }): void {
    const pick = wallsFor(this.box, eye);
    this.floor.visible = pick.floor;
    this.xMin.visible = pick.x === "min";
    this.xMax.visible = pick.x === "max";
    this.yMin.visible = pick.y === "min";
    this.yMax.visible = pick.y === "max";
  }

  /**
   * Keep each formation label a constant size on screen, sitting just above its dashed top.
   * `viewHeight` is the canvas height in CSS pixels and `fovDeg` is the camera's vertical field.
   */
  fitLabels(eye: { x: number; y: number; z: number }, fovDeg: number, viewHeight: number): void {
    if (!(viewHeight > 0) || this.labels.length === 0) return;
    const half = Math.tan((fovDeg * Math.PI) / 180 / 2);
    for (const label of this.labels) {
      const probe = this.place(label.wall, label.along, label.z, true);
      const distance = Math.max(1, Math.hypot(eye.x - probe[0], eye.y - probe[1], eye.z - probe[2]));
      const worldPerPx = (2 * half * distance) / viewHeight;
      const worldH = LABEL_SCREEN_PX * worldPerPx;
      const worldW = worldH * (label.pxWidth / label.pxHeight);
      const span = label.wall.alongMax - label.wall.alongMin;
      const widthFraction = span > 0 ? worldW / span : 0;
      let along = label.along;
      if (along + widthFraction > 0.98) along = Math.max(0.02, 0.98 - widthFraction);
      // The plane is centred on its position. Shift it by half its width so the name starts at
      // `along` and runs to the right, and by half its height so it sits on the dashed top.
      const start = this.place(label.wall, along, label.z, true);
      label.mesh.scale.set(worldW, worldH, 1);
      label.mesh.position.set(
        start[0] + label.wall.right[0] * worldW / 2,
        start[1] + label.wall.right[1] * worldW / 2,
        start[2] + worldH / 2,
      );
    }
  }
}
