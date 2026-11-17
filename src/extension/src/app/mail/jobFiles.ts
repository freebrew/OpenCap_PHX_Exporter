import type { FileEntry, FolderEntry } from "../pickFile";

async function openInside(root: FolderEntry, name: string): Promise<FileEntry | null> {
  const parts = name.split(/[\\/]/).filter((part) => part !== "");
  if (parts.length === 0 || parts.some((part) => part === "." || part === "..")) return null;
  try {
    let folder = root;
    for (const part of parts.slice(0, -1)) folder = await folder.getDirectoryHandle(part);
    const leaf = parts[parts.length - 1];
    if (!leaf) return null;
    return await folder.getFileHandle(leaf);
  } catch {
    return null;
  }
}

export async function readNamedFiles(root: FolderEntry | null, names: string[]): Promise<Array<{ name: string; bytes: Uint8Array | null }>> {
  const files: Array<{ name: string; bytes: Uint8Array | null }> = [];
  for (const name of names) {
    const handle = root ? await openInside(root, name) : null;
    files.push(handle ? { name, bytes: new Uint8Array(await (await handle.getFile()).arrayBuffer()) } : { name, bytes: null });
  }
  return files;
}
