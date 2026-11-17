import type { FolderEntry } from "../pickFile";

export function archiveFileName(day: Date): string {
  const month = String(day.getMonth() + 1).padStart(2, "0");
  const date = String(day.getDate()).padStart(2, "0");
  return `${day.getFullYear()}-${month}-${date} Daily Report.eml`;
}

export async function writeArchive(root: FolderEntry, eml: string, day = new Date()): Promise<string> {
  const folder = await root.getDirectoryHandle("Midnight Reports", { create: true });
  const name = archiveFileName(day);
  const handle = await folder.getFileHandle(name, { create: true });
  const writer = await handle.createWritable();
  await writer.write(eml);
  await writer.close();
  return name;
}
