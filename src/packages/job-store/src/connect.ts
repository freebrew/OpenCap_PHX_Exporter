import { directoryFromHandle, type DirectoryHandleLike, type StoredDirectory } from "./directory.js";
import type { JobRegistry } from "./registry.js";

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
