export const ODATA_ENTITIES = [
  "ClientJobs",
  "ClientJobCustomValues",
  "JobSchedules",
  "ToolAssemblies",
  "ToolAssemblyItems",
  "JobTools",
  "ActivityLogs",
  "SurveySheetEntries",
] as const;

export interface FieldCapRow {
  entity: string;
  values: Record<string, string>;
}

export function odataRows(payload: unknown): Record<string, unknown>[] {
  if (!payload || typeof payload !== "object") return [];
  const value = (payload as { value?: unknown }).value;
  return Array.isArray(value) ? value.filter((row): row is Record<string, unknown> => !!row && typeof row === "object") : [];
}

export function normalizeEntity(entity: string, payload: unknown): FieldCapRow[] {
  return odataRows(payload).map((row) => ({
    entity,
    values: Object.fromEntries(Object.entries(row).filter(([key]) => !key.startsWith("@")).map(([key, value]) => [key, value == null ? "" : String(value)])),
  }));
}

export function headersMatch(existing: string[], incoming: string[]): boolean {
  if (existing.length === 0) return false;
  return incoming.every((header) => existing.includes(header));
}
