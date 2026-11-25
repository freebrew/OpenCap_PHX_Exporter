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

/** Left drag orbits. A pan drag (right button or shift) slides the target. Phi stays off the poles so the view cannot tumble through them. */
export function dragOrbit(orbit: Orbit, dx: number, dy: number, pan: boolean): Orbit {
  if (!pan) {
    return {
      ...orbit,
      theta: orbit.theta - dx * 0.005,
      phi: Math.min(Math.PI - 0.12, Math.max(0.12, orbit.phi - dy * 0.005)),
    };
  }
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

function nearestOnPath(xyz: ArrayLike<number>, x: number, y: number, z: number): { x: number; y: number; z: number; d2: number } | null {
  const count = Math.floor(xyz.length / 3);
  if (count < 1) return null;
  let bestX = xyz[0]!;
  let bestY = xyz[1]!;
  let bestZ = xyz[2]!;
  let best = Infinity;
  const consider = (px: number, py: number, pz: number) => {
    const d2 = (px - x) ** 2 + (py - y) ** 2 + (pz - z) ** 2;
    if (d2 < best) {
      best = d2;
      bestX = px;
      bestY = py;
      bestZ = pz;
    }
  };
  if (count === 1) consider(bestX, bestY, bestZ);
  for (let i = 0; i < count - 1; i += 1) {
    const ax = xyz[i * 3]!;
    const ay = xyz[i * 3 + 1]!;
    const az = xyz[i * 3 + 2]!;
    const bx = xyz[(i + 1) * 3]!;
    const by = xyz[(i + 1) * 3 + 1]!;
    const bz = xyz[(i + 1) * 3 + 2]!;
    const abx = bx - ax;
    const aby = by - ay;
    const abz = bz - az;
    const ab2 = abx * abx + aby * aby + abz * abz;
    let t = 0;
    if (ab2 > 1e-8) {
      t = ((x - ax) * abx + (y - ay) * aby + (z - az) * abz) / ab2;
      if (t < 0) t = 0;
      else if (t > 1) t = 1;
    }
    consider(ax + abx * t, ay + aby * t, az + abz * t);
  }
  return { x: bestX, y: bestY, z: bestZ, d2: best };
}

/**
 * Hold the look-at within `leash` metres of the active well. Sliding along the well is free.
 * Sliding away stops at the leash, so a pan cannot walk the well out of the frame.
 * A leash of 0 sits the look-at on the well.
 */
export function leashOrbit(orbit: Orbit, paths: readonly ArrayLike<number>[], leash: number): Orbit {
  let near: { x: number; y: number; z: number; d2: number } | null = null;
  for (const path of paths) {
    const hit = nearestOnPath(path, orbit.targetX, orbit.targetY, orbit.targetZ);
    if (hit && (!near || hit.d2 < near.d2)) near = hit;
  }
  if (!near) return orbit;
  const dist = Math.sqrt(near.d2);
  if (dist <= leash) return orbit;
  const scale = leash <= 0 || dist === 0 ? 0 : leash / dist;
  return {
    ...orbit,
    targetX: near.x + (orbit.targetX - near.x) * scale,
    targetY: near.y + (orbit.targetY - near.y) * scale,
    targetZ: near.z + (orbit.targetZ - near.z) * scale,
  };
}
