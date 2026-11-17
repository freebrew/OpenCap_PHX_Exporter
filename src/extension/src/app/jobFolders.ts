// What a job folder holds for the preview, and the app's id for it.

export interface FolderWorkbooks {
  /** The saved snapshot (<workbook name>.phx.json). The preview writes it on Continue; nobody makes it by hand. */
  snapshot: string | null;
  /** The Excel workbook a snapshot is imported from. */
  xlsm: string | null;
}

const isSnapshot = (name: string) => name.endsWith(".phx.json") && !name.includes("backup");

/** "~$Slide Sheet.xlsm" is Excel's lock file while the workbook is open there, not a workbook. */
const isWorkbook = (name: string) => name.toLowerCase().endsWith(".xlsm") && !name.startsWith("~$");

/** The snapshot and workbook among a folder's file names. `expected` is the snapshot the job list has on record. */
export function folderWorkbooks(names: readonly string[], expected = ""): FolderWorkbooks {
  return {
    snapshot: (expected !== "" && names.includes(expected) ? expected : names.find(isSnapshot)) ?? null,
    xlsm: names.find(isWorkbook) ?? null,
  };
}

export interface ListedFolder {
  jobId: string;
  /** The snapshot this job-list entry opens. */
  fileName: string;
  /** The folder grant (a FileSystemDirectoryHandle). */
  handle?: unknown;
}

/**
 * Job-list entries that are a second listing of the folder being opened: same folder, and the
 * snapshot they expect is not in it. That is the entry made while the folder was still empty,
 * left behind when the folder was granted again under its workbook's name.
 */
export async function staleEntriesFor(entries: readonly ListedFolder[], keptId: string, handle: unknown, names: readonly string[]): Promise<string[]> {
  const opened = handle as { isSameEntry?: (other: unknown) => Promise<boolean> } | null | undefined;
  if (typeof opened?.isSameEntry !== "function") return [];
  const sameFolder = async (other: unknown) => {
    try {
      return await opened.isSameEntry!(other);
    } catch {
      return false;
    }
  };
  const stale: string[] = [];
  for (const entry of entries) {
    if (entry.jobId === keptId || !entry.handle || names.includes(entry.fileName)) continue;
    if (await sameFolder(entry.handle)) stale.push(entry.jobId);
  }
  return stale;
}

export function snapshotLabel(fileName: string): string {
  return fileName.replace(/\.phx\.json$/, "");
}

/** The job list's key for a folder, from the name shown for it. */
export function appJobIdOf(label: string): string {
  return label.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "") || "job";
}
