/** The parts of a three.js object the well renderer touches. three 0.187 ships no types, so this is the shape it relies on. */
export interface AnyObject {
  visible: boolean;
  renderOrder: number;
  position: { set: (x: number, y: number, z: number) => void };
  scale: { set: (x: number, y: number, z: number) => void; setScalar: (n: number) => void };
  quaternion: { setFromUnitVectors: (a: unknown, b: unknown) => void };
  updateMatrix: () => void;
  matrix: unknown;
  add: (child: AnyObject) => void;
  remove: (child: AnyObject) => void;
  children: AnyObject[];
  geometry?: { type?: string; dispose?: () => void };
  material?: { dispose?: () => void } | Array<{ dispose?: () => void }>;
}

/** A wide line through the points, in a colour, at a pixel width. */
export type StrokeMaker = (xyz: Float32Array, color: string, width: number, opacity: number) => AnyObject;
