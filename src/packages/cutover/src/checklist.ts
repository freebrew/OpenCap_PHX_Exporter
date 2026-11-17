export const CUTOVER_CHECKS = [
  { id: "calc", label: "Directional cells calculate in this preview." },
  { id: "picture", label: "The corridor picture uses this report day only." },
  { id: "report", label: "The morning report and .eml were reviewed." },
  { id: "sync", label: "FieldCap rows were reviewed before applying them." },
  { id: "fallback", label: "The Excel workbook stays on disk as the locked fallback. This preview never writes it." },
] as const;

export type CutoverId = (typeof CUTOVER_CHECKS)[number]["id"];

export function readyForNative(checked: Iterable<string>): boolean {
  const have = new Set(checked);
  return CUTOVER_CHECKS.every((item) => have.has(item.id));
}

export function authorityOf(sidecar: { authority?: unknown } | null | undefined): "excel" | "native" {
  return sidecar?.authority === "native" ? "native" : "excel";
}
