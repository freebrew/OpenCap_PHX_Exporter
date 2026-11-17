import { auditBundle, commitPlan, type HandoffReceipt, parseReceipt } from "@phx/xlsx-bridge";
import type { FolderEntry } from "./pickFile";

async function nest(root: FolderEntry, segments: string[], create: boolean): Promise<FolderEntry | null> {
  let folder = root;
  for (const segment of segments) {
    try {
      folder = await folder.getDirectoryHandle(segment, { create });
    } catch {
      return null;
    }
  }
  return folder;
}

async function writeText(folder: FolderEntry, name: string, text: string): Promise<void> {
  const handle = await folder.getFileHandle(name, { create: true });
  const writer = await handle.createWritable();
  await writer.write(text);
  await writer.close();
}

async function readText(folder: FolderEntry, name: string): Promise<string | null> {
  try {
    const handle = await folder.getFileHandle(name);
    return await (await handle.getFile()).text();
  } catch {
    return null;
  }
}

/** Payload files are read back before READY is written. Excel ignores a folder without READY. */
export async function commitHandoff(root: FolderEntry, id: string, files: Record<string, string>): Promise<void> {
  const plan = commitPlan(files);
  const dir = await nest(root, [".phx", "handoff", "pending", id], true);
  if (!dir) throw new Error("The handoff folder could not be created.");
  const readBack: Record<string, string> = {};
  for (const file of plan.payloads) {
    await writeText(dir, file.name, file.text);
    const text = await readText(dir, file.name);
    if (text !== file.text) throw new Error(`${file.name} did not read back from the job folder.`);
    readBack[file.name] = text;
  }
  const problems = await auditBundle(readBack);
  if (problems.length > 0) throw new Error(problems.join(" "));
  await writeText(dir, plan.ready.name, plan.ready.text);
}

export async function pendingHandoffs(root: FolderEntry): Promise<string[]> {
  const pending = await nest(root, [".phx", "handoff", "pending"], false);
  if (!pending) return [];
  const ready: string[] = [];
  for await (const [name, handle] of pending.entries()) {
    if (handle.kind !== "directory") continue;
    const marker = await readText(handle as FolderEntry, "READY");
    if (marker != null) ready.push(name);
  }
  return ready.sort();
}

export async function latestReceipt(root: FolderEntry): Promise<HandoffReceipt | null> {
  const receipts = await nest(root, [".phx", "handoff", "receipts"], false);
  if (!receipts) return null;
  let newest: HandoffReceipt | null = null;
  let newestStamp = "";
  for await (const [name, handle] of receipts.entries()) {
    if (handle.kind !== "file" || !name.endsWith(".csv")) continue;
    const text = await readText(receipts, name);
    if (!text) continue;
    const receipt = parseReceipt(text);
    if (name > newestStamp) {
      newestStamp = name;
      newest = receipt;
    }
  }
  return newest;
}

export function workbookToken(fileName: string): string {
  const digits = fileName.match(/(\d{4,})/);
  if (digits) return digits[1];
  const words = fileName.replace(/\.xlsm$/i, "").split(/[^A-Za-z0-9]+/).filter((part) => part.length > 0);
  return words[words.length - 1] ?? "";
}
