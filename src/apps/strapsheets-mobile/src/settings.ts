// The choices kept on the phone: which job, who receives the PDF, and whether several
// sheets go out as one file.

import { readJson, writeJson } from "./storage";

export interface Settings {
  jobNumber: string;
  /** Jobs used on this phone, newest first. */
  recentJobs: string[];
  recipients: string[];
  /** Several sessions in one PDF, or one PDF each. */
  combined: boolean;
}

const KEY = "settings";
const RECENT_JOBS = 10;

export function defaultSettings(): Settings {
  return { jobNumber: "", recentJobs: [], recipients: [], combined: true };
}

/** Keep the local choices from a saved blob and ignore anything left from an older sign-in build. */
export function settingsFromStored(stored: Partial<Settings> | null | undefined): Settings {
  const source = stored ?? {};
  return {
    jobNumber: typeof source.jobNumber === "string" ? source.jobNumber : "",
    recentJobs: stringList(source.recentJobs).slice(0, RECENT_JOBS),
    recipients: stringList(source.recipients),
    combined: source.combined !== false,
  };
}

export async function loadSettings(): Promise<Settings> {
  return settingsFromStored(await readJson<Partial<Settings> | null>(KEY, null));
}

export async function saveSettings(settings: Settings): Promise<void> {
  await writeJson(KEY, settingsFromStored(settings));
}

function stringList(value: unknown): string[] {
  if (!Array.isArray(value)) return [];
  return value.filter((item): item is string => typeof item === "string" && item.trim() !== "");
}
