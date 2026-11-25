import type { StoredDirectory } from "./directory.js";
import { StoreError } from "./types.js";

/** Single restore-point location under the job folder, next to the OpenCap CSV exports. */
export const BACKUP_DIR = "OpenCap";
export const BACKUP_FILE = ".phx-backup.json";

function isLegacyBackupName(name: string): boolean {
  return name.startsWith(".phx-backup-") && name.endsWith(".json");
}

function isBackupFileName(name: string): boolean {
  return name === BACKUP_FILE || isLegacyBackupName(name);
}

function compareLegacyNewestFirst(a: string, b: string): number {
  const rank = (name: string) => {
    if (name.includes("-recent-")) return 0;
    if (name.includes("-daily-")) return 1;
    if (name.includes("-weekly-")) return 2;
    return 3;
  };
  return rank(a) - rank(b) || b.localeCompare(a);
}

async function writeOverwrite(directory: StoredDirectory, name: string, bytes: Uint8Array): Promise<void> {
  const file = await directory.getFile(name, true);
  if (!file) throw new StoreError("invalid", `Backup file ${name} could not be created.`);
  const writer = await file.writeExclusive();
  try {
    await writer.write(bytes);
    await writer.close();
  } catch (error) {
    await writer.abort().catch(() => undefined);
    throw error;
  }
}

async function removeFiles(directory: StoredDirectory, names: readonly string[]): Promise<void> {
  for (const name of names) {
    const file = await directory.getFile(name, false);
    if (file) await file.remove().catch(() => undefined);
  }
}

/**
 * Overwrite the single restore point at `<job>/OpenCap/.phx-backup.json` and remove any
 * legacy multi-backup files from the job root or `OpenCap/`.
 */
export async function createRestorePoint(directory: StoredDirectory, canonical: Uint8Array): Promise<void> {
  const backupDir = await directory.getDirectory(BACKUP_DIR, true);
  if (!backupDir) throw new StoreError("invalid", `${BACKUP_DIR} folder could not be created for backup.`);
  await writeOverwrite(backupDir, BACKUP_FILE, canonical);
  const [rootNames, openCapNames] = await Promise.all([directory.list(), backupDir.list()]);
  await removeFiles(directory, rootNames.filter(isBackupFileName));
  await removeFiles(backupDir, openCapNames.filter(isLegacyBackupName));
}

/**
 * Consolidate legacy `.phx-backup-*.json` files left in the job root or `OpenCap/` by older builds
 * into the single `OpenCap/.phx-backup.json` file, preserving the newest backup if none exists yet.
 */
export async function migrateLegacyBackups(directory: StoredDirectory): Promise<void> {
  const rootNames = await directory.list();
  const rootBackups = rootNames.filter(isBackupFileName);
  const openCap = await directory.getDirectory(BACKUP_DIR, false);
  const openCapNames = openCap ? await openCap.list() : [];
  const openCapLegacy = openCapNames.filter(isLegacyBackupName);
  if (rootBackups.length === 0 && openCapLegacy.length === 0) return;

  const backupDir = openCap ?? (await directory.getDirectory(BACKUP_DIR, true));
  if (!backupDir) return;

  const existingFile = await backupDir.getFile(BACKUP_FILE, false);
  const existingBytes = existingFile ? await existingFile.read() : null;
  if (!existingBytes || existingBytes.byteLength === 0) {
    const candidates: Array<{ dir: StoredDirectory; name: string }> = [
      ...openCapLegacy.map((name) => ({ dir: backupDir, name })),
      ...rootBackups.map((name) => ({ dir: directory, name })),
    ].sort((a, b) => compareLegacyNewestFirst(a.name, b.name));
    for (const candidate of candidates) {
      const file = await candidate.dir.getFile(candidate.name, false);
      const bytes = file ? await file.read() : null;
      if (bytes && bytes.byteLength > 0) {
        await writeOverwrite(backupDir, BACKUP_FILE, bytes);
        break;
      }
    }
  }

  await removeFiles(directory, rootBackups);
  await removeFiles(backupDir, openCapLegacy);
}

/** Available restore points: `OpenCap/.phx-backup.json` (plus any un-migrated legacy backups). */
export async function listBackups(directory: StoredDirectory): Promise<string[]> {
  const found: string[] = [];
  const openCap = await directory.getDirectory(BACKUP_DIR, false);
  if (openCap) {
    for (const name of (await openCap.list()).filter(isBackupFileName).sort()) {
      found.push(`${BACKUP_DIR}/${name}`);
    }
  }
  for (const name of (await directory.list()).filter(isBackupFileName).sort()) {
    found.push(name);
  }
  return found;
}
