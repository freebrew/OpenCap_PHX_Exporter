// Orbit with Z up. theta swings around up, phi is the angle down from up.

export interface Orbit {
  targetX: number;
  targetY: number;
  targetZ: number;
  theta: number;
  phi: number;
  radius: number;
}

export interface Bounds {
  minX: number;
  maxX: number;
  minY: number;
  maxY: number;
  minZ: number;
  maxZ: number;
}

export function emptyBounds(): Bounds {
  return { minX: Infinity, maxX: -Infinity, minY: Infinity, maxY: -Infinity, minZ: Infinity, maxZ: -Infinity };
}

export function includeXyz(bounds: Bounds, xyz: ArrayLike<number>): void {
  for (let i = 0; i < xyz.length; i += 3) {
    const x = xyz[i]!;
    const y = xyz[i + 1]!;
    const z = xyz[i + 2]!;
    if (x < bounds.minX) bounds.minX = x;
    if (x > bounds.maxX) bounds.maxX = x;
    if (y < bounds.minY) bounds.minY = y;
    if (y > bounds.maxY) bounds.maxY = y;
    if (z < bounds.minZ) bounds.minZ = z;
    if (z > bounds.maxZ) bounds.maxZ = z;
  }
}

export function fitOrbit(bounds: Bounds): Orbit {
  const spanX = Number.isFinite(bounds.minX) ? bounds.maxX - bounds.minX : 0;
  const spanY = Number.isFinite(bounds.minY) ? bounds.maxY - bounds.minY : 0;
  const spanZ = Number.isFinite(bounds.minZ) ? bounds.maxZ - bounds.minZ : 0;
  const span = Math.max(spanX, spanY, spanZ, 1);
  return {
    targetX: Number.isFinite(bounds.minX) ? (bounds.minX + bounds.maxX) / 2 : 0,
    targetY: Number.isFinite(bounds.minY) ? (bounds.minY + bounds.maxY) / 2 : 0,
    targetZ: Number.isFinite(bounds.minZ) ? (bounds.minZ + bounds.maxZ) / 2 : 0,
    theta: -Math.PI / 4,
    phi: Math.PI / 3,
    radius: span * 1.8,
  };
}

export function cameraFromOrbit(orbit: Orbit): { x: number; y: number; z: number } {
  const phi = Math.min(Math.PI - 0.08, Math.max(0.08, orbit.phi));
  const radius = Math.min(1_000_000, Math.max(0.5, orbit.radius));
  return {
    x: orbit.targetX + radius * Math.sin(phi) * Math.cos(orbit.theta),
    y: orbit.targetY + radius * Math.sin(phi) * Math.sin(orbit.theta),
    z: orbit.targetZ + radius * Math.cos(phi),
  };
}

/** Left drag orbits. A pan drag (right button or shift) slides the target. */
export function dragOrbit(orbit: Orbit, dx: number, dy: number, pan: boolean): Orbit {
  if (!pan) return { ...orbit, theta: orbit.theta - dx * 0.005, phi: orbit.phi - dy * 0.005 };
  const eye = cameraFromOrbit(orbit);
  let fx = orbit.targetX - eye.x;
  let fy = orbit.targetY - eye.y;
  let fz = orbit.targetZ - eye.z;
  const length = Math.hypot(fx, fy, fz) || 1;
  fx /= length;
  fy /= length;
  fz /= length;
  let rx = fy;
  let ry = -fx;
  const right = Math.hypot(rx, ry) || 1;
  rx /= right;
  ry /= right;
  const ux = ry * fz;
  const uy = -rx * fz;
  const uz = rx * fy - ry * fx;
  const scale = Math.max(0.5, orbit.radius) * 0.002;
  return {
    ...orbit,
    targetX: orbit.targetX - rx * dx * scale + ux * dy * scale,
    targetY: orbit.targetY - ry * dx * scale + uy * dy * scale,
    targetZ: orbit.targetZ + uz * dy * scale,
  };
}

export function wheelOrbit(orbit: Orbit, deltaY: number): Orbit {
  return { ...orbit, radius: Math.min(1_000_000, Math.max(1, orbit.radius * Math.exp(deltaY * 0.001))) };
}
