import { createRestorePoint, listBackups } from "./backups.js";
import type { StoredDirectory, StoredFile } from "./directory.js";
import { packMirror } from "./mirror.js";
import type { JobRegistry } from "./registry.js";
import { encodeEnvelope, parseEnvelope, semanticSignature, sha256 } from "./snapshot.js";
import { StoreError, type PhxEnvelope, type Recovery, type SaveState } from "./types.js";

export interface OpenResult {
  envelope: PhxEnvelope | null;
  recovery: Recovery;
  state: SaveState;
}

async function verifiedWrite(file: StoredFile, bytes: Uint8Array, previous: Uint8Array | null): Promise<void> {
  const writer = await file.writeExclusive();
  try {
    await writer.write(bytes);
    await writer.close();
  } catch (error) {
    await writer.abort().catch(() => undefined);
    throw error;
  }
  const readBack = await file.read();
  if (!readBack) throw new StoreError("verify", "Snapshot disappeared after write.");
  const [written, actual] = await Promise.all([sha256(bytes), sha256(readBack)]);
  if (written !== actual) {
    await restore(file, previous);
    throw new StoreError("verify", "Snapshot verification failed and the previous copy was restored.");
  }
  let parsed: PhxEnvelope;
  try {
    parsed = parseEnvelope(readBack);
  } catch (error) {
    await restore(file, previous);
    throw error;
  }
  if (semanticSignature(parsed) !== semanticSignature(parseEnvelope(bytes))) {
    await restore(file, previous);
    throw new StoreError("verify", "Snapshot semantic check failed and the previous copy was restored.");
  }
}

async function restore(file: StoredFile, previous: Uint8Array | null): Promise<void> {
  if (!previous || previous.byteLength === 0) return;
  const writer = await file.writeExclusive();
  await writer.write(previous);
  await writer.close();
}

export class JobSession {
  state: SaveState = "folder-safe";
  /** Why the last browser safety copy was not written, or "" when it was. */
  mirrorError = "";
  private mirrorTimer: ReturnType<typeof setTimeout> | null = null;
  private folderTimer: ReturnType<typeof setTimeout> | null = null;
  private canonicalName: string;

  constructor(
    readonly jobId: string,
    readonly label: string,
    private readonly directory: StoredDirectory,
    private readonly registry: JobRegistry,
    fileName: string,
    private envelope: PhxEnvelope | null,
  ) {
    this.canonicalName = fileName;
  }

  get current(): PhxEnvelope | null {
    return this.envelope;
  }

  get fileName(): string {
    return this.canonicalName;
  }

  dirty(): boolean {
    return this.state === "dirty" || this.state === "saving-mirror" || this.state === "mirror-safe" || this.state === "save-error";
  }

  folderUnsafe(): boolean {
    return this.state !== "folder-safe";
  }

  async replace(envelope: PhxEnvelope, options: { backup: boolean }): Promise<void> {
    this.envelope = envelope;
    this.state = "dirty";
    await this.writeMirror();
    await this.writeFolder(options.backup);
  }

  hold(envelope: PhxEnvelope | null): void {
    this.envelope = envelope;
  }

  noteEdit(envelope: PhxEnvelope): void {
    this.envelope = envelope;
    this.state = "dirty";
    if (this.mirrorTimer) clearTimeout(this.mirrorTimer);
    this.mirrorTimer = setTimeout(() => {
      void this.writeMirror();
    }, 2_000);
    if (this.folderTimer) clearTimeout(this.folderTimer);
    this.folderTimer = setTimeout(() => {
      void this.writeFolder(false);
    }, 30_000);
  }

  async flushFolder(): Promise<void> {
    if (this.mirrorTimer) clearTimeout(this.mirrorTimer);
    if (this.folderTimer) clearTimeout(this.folderTimer);
    await this.writeMirror();
    await this.writeFolder(true);
  }

  /**
   * Never throws. The folder file is the record and the mirror only covers the gap before it is
   * written, so a mirror that cannot be stored must not stop the folder save that follows it.
   */
  async writeMirror(): Promise<void> {
    if (!this.envelope) return;
    this.state = "saving-mirror";
    try {
      await this.registry.putMirror(await packMirror(this.jobId, this.envelope.savedAt, encodeEnvelope(this.envelope)));
      this.mirrorError = "";
      if (this.state === "saving-mirror") this.state = "mirror-safe";
    } catch (error) {
      this.mirrorError = error instanceof Error ? error.message : String(error);
      if (this.state === "saving-mirror") this.state = "dirty";
    }
  }

  async writeFolder(backup: boolean): Promise<void> {
    if (!this.envelope) return;
    const permission = await this.directory.queryPermission();
    if (permission !== "granted") {
      this.state = "permission-blocked";
      return;
    }
    this.state = "saving-folder";
    const file = await this.directory.getFile(this.canonicalName, true);
    if (!file) throw new StoreError("invalid", "Snapshot file could not be created.");
    const previous = await file.read();
    if (backup && previous && previous.byteLength > 0) await createRestorePoint(this.directory, previous);
    const bytes = encodeEnvelope(this.envelope);
    try {
      await verifiedWrite(file, bytes, previous && previous.byteLength > 0 ? previous : null);
      this.state = "folder-safe";
    } catch (error) {
      this.state = error instanceof StoreError && error.code === "concurrent" ? "concurrent-writer" : "save-error";
      throw error;
    }
  }

  dispose(): void {
    if (this.mirrorTimer) clearTimeout(this.mirrorTimer);
    if (this.folderTimer) clearTimeout(this.folderTimer);
  }
}

export async function openJob(input: {
  jobId: string;
  label: string;
  fileName: string;
  directory: StoredDirectory;
  registry: JobRegistry;
}): Promise<{ session: JobSession; result: OpenResult }> {
  const permission = await input.directory.queryPermission();
  const session = new JobSession(input.jobId, input.label, input.directory, input.registry, input.fileName, null);
  if (permission !== "granted") {
    session.state = "permission-blocked";
    return {
      session,
      result: {
        envelope: null,
        recovery: { kind: "permission", jobId: input.jobId, label: input.label },
        state: session.state,
      },
    };
  }
  const file = await input.directory.getFile(input.fileName, false);
  const bytes = file ? await file.read() : null;
  const mirror = await input.registry.getMirror(input.jobId);
  if (!bytes || bytes.byteLength === 0) {
    return { session, result: { envelope: null, recovery: { kind: "none" }, state: "folder-safe" } };
  }
  try {
    const envelope = parseEnvelope(bytes);
    session.hold(envelope);
    if (mirror && Date.parse(mirror.savedAt) > Date.parse(envelope.savedAt)) {
      return {
        session,
        result: {
          envelope,
          recovery: {
            kind: "mirror-newer",
            jobId: input.jobId,
            label: input.label,
            mirrorSavedAt: mirror.savedAt,
            folderSavedAt: envelope.savedAt,
          },
          state: "folder-safe",
        },
      };
    }
    return { session, result: { envelope, recovery: { kind: "none" }, state: "folder-safe" } };
  } catch (error) {
    const backups = await listBackups(input.directory);
    const kind = error instanceof StoreError && error.code === "future" ? "future" : "corrupt";
    return {
      session,
      result: {
        envelope: null,
        recovery: { kind, jobId: input.jobId, label: input.label, backups },
        state: "save-error",
      },
    };
  }
}
