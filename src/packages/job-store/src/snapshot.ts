import { SCHEMA_VERSION, StoreError, type PhxEnvelope } from "./types.js";

export function createEnvelope<TWorkbook, TSidecar>(input: {
  univerAppVersion: string;
  sourceWorkbookHash: string;
  workbook: TWorkbook;
  sidecar: TSidecar;
  savedAt?: string;
  importBaseline?: unknown;
  handoffBlocks?: string[];
}): PhxEnvelope<TWorkbook, TSidecar> {
  return {
    schemaVersion: SCHEMA_VERSION,
    univerAppVersion: input.univerAppVersion,
    sourceWorkbookHash: input.sourceWorkbookHash,
    savedAt: input.savedAt ?? new Date().toISOString(),
    workbook: input.workbook,
    sidecar: input.sidecar,
    ...(input.importBaseline != null ? { importBaseline: input.importBaseline } : {}),
    ...(input.handoffBlocks && input.handoffBlocks.length > 0 ? { handoffBlocks: input.handoffBlocks } : {}),
  };
}

export function encodeEnvelope(envelope: PhxEnvelope): Uint8Array {
  return new TextEncoder().encode(JSON.stringify(envelope));
}

export function parseEnvelope(bytes: Uint8Array): PhxEnvelope {
  let text: string;
  try {
    text = new TextDecoder().decode(bytes);
  } catch {
    throw new StoreError("corrupt", "Snapshot bytes could not be decoded.");
  }
  let value: unknown;
  try {
    value = JSON.parse(text);
  } catch {
    throw new StoreError("corrupt", "Snapshot is not valid JSON.");
  }
  if (!value || typeof value !== "object") throw new StoreError("corrupt", "Snapshot is empty.");
  const record = value as Partial<PhxEnvelope>;
  if (typeof record.schemaVersion !== "number") throw new StoreError("corrupt", "Snapshot has no schema version.");
  if (record.schemaVersion > SCHEMA_VERSION) {
    throw new StoreError("future", `Snapshot schema ${record.schemaVersion} is newer than this preview.`);
  }
  if (record.schemaVersion !== 1) throw new StoreError("invalid", "Snapshot schema is not supported.");
  if (!record.workbook || typeof record.workbook !== "object") throw new StoreError("corrupt", "Snapshot workbook is missing.");
  const baseline = (record as { importBaseline?: unknown }).importBaseline;
  const blocks = (record as { handoffBlocks?: unknown }).handoffBlocks;
  return {
    schemaVersion: 1,
    univerAppVersion: String(record.univerAppVersion ?? ""),
    sourceWorkbookHash: String(record.sourceWorkbookHash ?? ""),
    savedAt: String(record.savedAt ?? ""),
    workbook: preserveWorkbook(record.workbook),
    sidecar: record.sidecar ?? {},
    ...(baseline != null ? { importBaseline: baseline } : {}),
    ...(Array.isArray(blocks) ? { handoffBlocks: blocks.map((item) => String(item)) } : {}),
  };
}

function preserveWorkbook(workbook: object): object {
  return workbook;
}

export async function sha256(bytes: Uint8Array): Promise<string> {
  const digest = await crypto.subtle.digest("SHA-256", bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer);
  return [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, "0")).join("");
}

export function semanticSignature(envelope: PhxEnvelope): string {
  const workbook = envelope.workbook as { sheetOrder?: unknown; resources?: unknown; sheets?: unknown };
  return JSON.stringify({
    schemaVersion: envelope.schemaVersion,
    sourceWorkbookHash: envelope.sourceWorkbookHash,
    sheetOrder: workbook.sheetOrder ?? null,
    resources: workbook.resources ?? null,
    sidecar: envelope.sidecar,
    importBaseline: envelope.importBaseline ?? null,
    handoffBlocks: envelope.handoffBlocks ?? [],
  });
}
