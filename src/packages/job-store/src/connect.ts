import { directoryFromHandle, type StoredDirectory } from "./directory.js";
import type { JobRegistry } from "./registry.js";

interface DirectoryHandleLike {
  getFileHandle(name: string, options?: { create?: boolean }): Promise<{
    getFile(): Promise<Blob>;
    createWritable(options: { mode?: "exclusive" | "siloed" }): Promise<{
      write(data: BufferSource | Blob | string): Promise<void>;
      close(): Promise<void>;
      abort(): Promise<void>;
    }>;
  }>;
  entries(): AsyncIterable<[string, { kind: string }]>;
  queryPermission(descriptor: { mode: "readwrite" }): Promise<"granted" | "prompt" | "denied">;
  requestPermission(descriptor: { mode: "readwrite" }): Promise<"granted" | "prompt" | "denied">;
  removeEntry?(name: string): Promise<void>;
}

export async function grantJobFolder(
  requestHandle: () => Promise<DirectoryHandleLike>,
  registry: JobRegistry,
  meta: { jobId: string; label: string; fileName: string },
): Promise<StoredDirectory> {
  const handle = await requestHandle();
  await registry.put({ ...meta, lastOpened: new Date().toISOString(), handle });
  return directoryFromHandle(handle);
}

export async function reopenStoredFolder(registry: JobRegistry, jobId: string): Promise<StoredDirectory | null> {
  const record = await registry.get(jobId);
  const handle = record?.handle as DirectoryHandleLike | undefined;
  if (!handle || typeof handle.queryPermission !== "function") return null;
  return directoryFromHandle(handle);
}
