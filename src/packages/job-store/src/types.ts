export const SCHEMA_VERSION = 1;

export type SaveState =
  | "dirty"
  | "saving-mirror"
  | "mirror-safe"
  | "saving-folder"
  | "folder-safe"
  | "permission-blocked"
  | "concurrent-writer"
  | "save-error";

export interface PhxEnvelope<TWorkbook = unknown, TSidecar = unknown> {
  schemaVersion: number;
  univerAppVersion: string;
  sourceWorkbookHash: string;
  savedAt: string;
  workbook: TWorkbook;
  sidecar: TSidecar;
  /** Immutable canonical workbook captured when the browser shift started. */
  importBaseline?: unknown;
  /** Commands from this shift that a handoff cannot represent. */
  handoffBlocks?: string[];
}

export interface JobRecord {
  jobId: string;
  label: string;
  fileName: string;
  lastOpened: string;
}

/** Read and written through `packMirror` / `mirrorSnapshot`. */
export interface MirrorRecord {
  jobId: string;
  savedAt: string;
  /** The snapshot file's bytes, gzipped. */
  gzip?: Uint8Array;
  /** How previews up to 0.1.44 kept it: the snapshot as text. Still read, never written. */
  bytes?: string;
}

export type Recovery =
  | { kind: "none" }
  | { kind: "permission"; jobId: string; label: string }
  | { kind: "corrupt" | "future"; jobId: string; label: string; backups: string[] }
  | { kind: "mirror-newer"; jobId: string; label: string; mirrorSavedAt: string; folderSavedAt: string };

export class StoreError extends Error {
  readonly code: "corrupt" | "future" | "invalid" | "concurrent" | "permission" | "verify";

  constructor(code: StoreError["code"], message: string) {
    super(message);
    this.name = "StoreError";
    this.code = code;
  }
}
