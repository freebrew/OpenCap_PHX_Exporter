import type { StoredDirectory } from "./directory.js";

const RECENT_LIMIT = 10;
const DAILY_LIMIT = 14;
const WEEKLY_LIMIT = 8;
const DAILY_AGE_MS = 14 * 24 * 60 * 60 * 1000;
const WEEKLY_AGE_MS = 56 * 24 * 60 * 60 * 1000;

export interface BackupPlan {
  recent: string;
  daily: string;
  weekly: string;
}

export function backupNames(now: Date): BackupPlan {
  const stamp = now.toISOString().replace(/[:.]/g, "-");
  const day = now.toISOString().slice(0, 10);
  const week = isoWeek(now);
  return {
    recent: `.phx-backup-recent-${stamp}.json`,
    daily: `.phx-backup-daily-${day}.json`,
    weekly: `.phx-backup-weekly-${week}.json`,
  };
}

function isoWeek(date: Date): string {
  const utc = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
  const day = utc.getUTCDay() || 7;
  utc.setUTCDate(utc.getUTCDate() + 4 - day);
  const yearStart = new Date(Date.UTC(utc.getUTCFullYear(), 0, 1));
  const week = Math.ceil(((utc.getTime() - yearStart.getTime()) / 86400000 + 1) / 7);
  return `${utc.getUTCFullYear()}-W${String(week).padStart(2, "0")}`;
}

async function writeCopy(directory: StoredDirectory, name: string, bytes: Uint8Array): Promise<void> {
  const existing = await directory.getFile(name, false);
  if (existing) return;
  const file = await directory.getFile(name, true);
  if (!file) return;
  const writer = await file.writeExclusive();
  await writer.write(bytes);
  await writer.close();
}

export async function createRestorePoint(directory: StoredDirectory, canonical: Uint8Array, now = new Date()): Promise<void> {
  const names = backupNames(now);
  await writeCopy(directory, names.recent, canonical);
  await writeCopy(directory, names.daily, canonical);
  await writeCopy(directory, names.weekly, canonical);
  await retain(directory, now);
}

export async function listBackups(directory: StoredDirectory): Promise<string[]> {
  const names = await directory.list();
  return names.filter((name) => name.startsWith(".phx-backup-")).sort();
}

async function retain(directory: StoredDirectory, now: Date): Promise<void> {
  const names = await listBackups(directory);
  await trim(directory, names.filter((name) => name.includes("-recent-")), RECENT_LIMIT, null, now);
  await trim(directory, names.filter((name) => name.includes("-daily-")), DAILY_LIMIT, DAILY_AGE_MS, now);
  await trim(directory, names.filter((name) => name.includes("-weekly-")), WEEKLY_LIMIT, WEEKLY_AGE_MS, now);
}

async function trim(directory: StoredDirectory, names: string[], limit: number, maxAge: number | null, now: Date): Promise<void> {
  const ordered = [...names].sort();
  const extras = ordered.slice(0, Math.max(0, ordered.length - limit));
  for (const name of ordered) {
    const stale = maxAge != null && now.getTime() - stampOf(name) > maxAge;
    if (!extras.includes(name) && !stale) continue;
    const file = await directory.getFile(name, false);
    if (file) await file.remove();
  }
}

function stampOf(name: string): number {
  const match = name.match(/(\d{4}-\d{2}-\d{2})/);
  if (!match) return Date.now();
  return Date.parse(match[1]);
}
