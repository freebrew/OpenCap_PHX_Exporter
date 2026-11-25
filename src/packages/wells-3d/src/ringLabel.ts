/** Every target is the same two rings: 10 m inside 20 m. */
export function ringRadii(): { outer: number; inner: number } {
  return { outer: 20, inner: 10 };
}

/** Whole metres only: "140 m", never "135.7 m". */
export function ringMetresLabel(metres: number): string {
  if (!Number.isFinite(metres)) return "";
  return `${Math.round(metres).toLocaleString("en-US")} m`;
}

/**
 * The label sits just inside the ring, on the line from the centre out to `rim`.
 * `inset` is how far inside the stroke the text centre belongs, in metres.
 * A large inset stops short of the centre so the words stay on their own ring.
 */
export function ringLabelPoint(
  center: { x: number; y: number; z: number },
  rim: { x: number; y: number; z: number },
  inset: number,
): { x: number; y: number; z: number } {
  const dx = rim.x - center.x;
  const dy = rim.y - center.y;
  const dz = rim.z - center.z;
  const len = Math.hypot(dx, dy, dz);
  if (!(len > 1e-6)) return { x: center.x, y: center.y, z: center.z };
  const pull = Math.min(len * 0.85, Math.max(0, inset));
  return {
    x: rim.x - (dx / len) * pull,
    y: rim.y - (dy / len) * pull,
    z: rim.z - (dz / len) * pull,
  };
}
