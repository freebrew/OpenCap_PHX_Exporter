// OpenCap CSVs for Setup's REFRESH CSVs: the newest export per file in the job folder, or a fresh
// pull from FieldCap through the OpenCap exporter running in this extension (read-only on FieldCap).
import { OPEN_CAP_FILES, type OpenCapKind, type OpenCapSource } from "../../../packages/workbook-commands/src/index";
import type { FileEntry, FolderEntry } from "./pickFile";

export const OPENCAP_DIR = "OpenCap";

/** chrome.storage.local keys the exporter fills after FETCH_ALL (only for CSVs that had rows). */
const EXPORTER_KEYS: Record<OpenCapKind, string> = {
  job: "fieldcap_csv_job",
  crew: "fieldcap_csv_crew",
  bha: "fieldcap_csv_bha",
  slide: "fieldcap_csv_slide_by_day",
  inventory: "fieldcap_csv_inventory",
  costs: "fieldcap_csv_ticket_costs",
};
const EXPORTER_META = "fieldcap_meta";
/** Per-activity breakdown behind the slide/rotate hours. Saved next to the CSVs; Excel does not load it. */
const EXPORTER_ACTIVITY_LOG = "fieldcap_csv_activity_log";
const ACTIVITY_LOG_SUFFIX = "activity-log-detail";

/** The exporter's download names, fieldcap-job-<id>-<suffix>.csv, so Excel's REFRESH CSVs finds the same files. */
const EXPORTER_SUFFIX: Record<OpenCapKind, string> = {
  job: "job-details",
  crew: "crew",
  bha: "bha-equipment",
  slide: "slide-rotate-metres-by-day",
  inventory: "inventory",
  costs: "ticket-costs-by-day",
};

const JOB_MAP_KEY = "phx_fieldcap_job_map";

export type OpenCapCsvs = Partial<Record<OpenCapKind, string>> & { activityLog?: string };

export interface FolderCsvs {
  sources: OpenCapSource[];
  /** What the files panel shows for a missing file. */
  folderPath: string;
}

interface Hit {
  file: File;
  path: string;
}

async function considerFolder(folder: FolderEntry, path: string, best: Map<OpenCapKind, Hit>): Promise<void> {
  for await (const [name, entry] of folder.entries()) {
    const lower = name.toLowerCase();
    if (entry.kind !== "file" || !lower.endsWith(".csv")) continue;
    const matches = OPEN_CAP_FILES.filter((spec) => lower.includes(spec.token));
    if (matches.length === 0) continue;
    const file = await (entry as FileEntry).getFile();
    for (const spec of matches) {
      const held = best.get(spec.kind);
      if (!held || file.lastModified >= held.file.lastModified) best.set(spec.kind, { file, path: `${path}\\${name}` });
    }
  }
}

/** FindCsvByToken: the newest *token*.csv in <job>\OpenCap, then the job folder itself (a tie goes to the later folder). */
export async function findFolderCsvs(root: FolderEntry): Promise<FolderCsvs> {
  const best = new Map<OpenCapKind, Hit>();
  const openCap = await root.getDirectoryHandle(OPENCAP_DIR).catch(() => null);
  if (openCap) await considerFolder(openCap, `${root.name}\\${OPENCAP_DIR}`, best);
  await considerFolder(root, root.name, best);
  const sources = await Promise.all([...best].map(async ([kind, hit]) => ({ kind, csv: await hit.file.text(), name: hit.file.name, path: hit.path })));
  return { sources, folderPath: `${root.name}\\${OPENCAP_DIR}` };
}

/** Save a pull into <job>\OpenCap under the exporter's names. Returns the sources as written. */
export async function writeFolderCsvs(root: FolderEntry, jobNumber: string, csvs: OpenCapCsvs): Promise<FolderCsvs> {
  const folder = await root.getDirectoryHandle(OPENCAP_DIR, { create: true });
  const sources: OpenCapSource[] = [];
  for (const spec of OPEN_CAP_FILES) {
    const csv = csvs[spec.kind];
    if (!csv) continue;
    const name = `fieldcap-job-${jobNumber}-${EXPORTER_SUFFIX[spec.kind]}.csv`;
    const writable = await (await folder.getFileHandle(name, { create: true })).createWritable();
    await writable.write(csv);
    await writable.close();
    sources.push({ kind: spec.kind, csv, name, path: `${root.name}\\${OPENCAP_DIR}\\${name}` });
  }
  if (csvs.activityLog) {
    const name = `fieldcap-job-${jobNumber}-${ACTIVITY_LOG_SUFFIX}.csv`;
    const writable = await (await folder.getFileHandle(name, { create: true })).createWritable();
    await writable.write(csvs.activityLog);
    await writable.close();
  }
  return { sources, folderPath: `${root.name}\\${OPENCAP_DIR}` };
}

export interface PullProgress {
  pct: number;
  label: string;
}

/** The exporter's FETCH_ALL for one FieldCap job: every file, CSV text only. */
export async function pullFromFieldCap(fieldcapJobId: string, onProgress: (progress: PullProgress) => void): Promise<OpenCapCsvs> {
  const keys = [...Object.values(EXPORTER_KEYS), EXPORTER_ACTIVITY_LOG];
  // The exporter stores only non-empty CSVs, so a file another job left behind must not survive this pull.
  await chrome.storage.local.remove([...keys, EXPORTER_META]);
  await new Promise<void>((resolve, reject) => {
    const port = chrome.runtime.connect({ name: "fieldcap-fetch" });
    let settled = false;
    port.onMessage.addListener((message: { type?: string; pct?: number; label?: string; ok?: boolean; error?: string }) => {
      if (message.type === "PROGRESS") onProgress({ pct: Number(message.pct ?? 0), label: String(message.label ?? "") });
      if (message.type !== "DONE") return;
      settled = true;
      port.disconnect();
      if (message.ok) resolve();
      else reject(new Error(message.error || "FieldCap pull failed."));
    });
    port.onDisconnect.addListener(() => {
      if (!settled) reject(new Error("The FieldCap pull stopped before it finished. Reload the extension and try again."));
    });
    port.postMessage({ type: "FETCH_ALL", jobId: fieldcapJobId, flags: { jobDetails: true, crew: true, bha: true, inventory: true, ticketCosts: true } });
  });
  const stored = await chrome.storage.local.get([...keys, EXPORTER_META]);
  const meta = stored[EXPORTER_META] as { jobId?: unknown } | undefined;
  if (String(meta?.jobId ?? "") !== fieldcapJobId) throw new Error(`The pull finished for job ${String(meta?.jobId ?? "?")}, not ${fieldcapJobId}. Nothing was loaded.`);
  const csvs: OpenCapCsvs = {};
  for (const spec of OPEN_CAP_FILES) {
    const text = stored[EXPORTER_KEYS[spec.kind]];
    if (typeof text === "string" && text !== "") csvs[spec.kind] = text;
  }
  const activityLog = stored[EXPORTER_ACTIVITY_LOG];
  if (typeof activityLog === "string" && activityLog !== "") csvs.activityLog = activityLog;
  return csvs;
}

/** Job Number to app job folder. */
export async function jobMap(): Promise<Record<string, string>> {
  if (!globalThis.chrome?.storage) return {};
  const stored = await chrome.storage.local.get(JOB_MAP_KEY);
  const map = stored[JOB_MAP_KEY];
  return map && typeof map === "object" ? map as Record<string, string> : {};
}

/** Remember which app job folder holds a Job Number. */
export async function rememberJob(jobNumber: string, appJobId: string): Promise<void> {
  await chrome.storage.local.set({ [JOB_MAP_KEY]: { ...(await jobMap()), [jobNumber]: appJobId } });
}

/** Two job-list entries for one folder became one: the Job Numbers held by the entry that went follow the one kept. */
export async function moveJobLinks(fromAppJobId: string, toAppJobId: string): Promise<void> {
  if (!globalThis.chrome?.storage) return;
  const moved = Object.entries(await jobMap()).map(([jobNumber, appJobId]) => [jobNumber, appJobId === fromAppJobId ? toAppJobId : appJobId]);
  await chrome.storage.local.set({ [JOB_MAP_KEY]: Object.fromEntries(moved) });
}

export interface FieldCapJob {
  /** FieldCap's internal id (the /Jobs/<id> in the URL). OData requests use this. */
  clientJobId: string;
  /** The Job Number people use (OData JobCode). */
  jobNumber: string;
  name: string;
  /** Field Active, Field Staging, Field Complete, ... */
  status: string;
  /** FieldCap has an end date ("To") for it. */
  ended: boolean;
}

const PULL_KEY = (jobNumber: string) => `phx_pull_${jobNumber}`;

export interface CachedPull {
  csvs: OpenCapCsvs;
  pulledAt: string;
}

/** Keep the latest pull of a job in the browser, so a job with no folder yet can still be published and filed later. */
export async function cachePull(jobNumber: string, csvs: OpenCapCsvs): Promise<void> {
  await chrome.storage.local.set({ [PULL_KEY(jobNumber)]: { csvs, pulledAt: new Date().toISOString() } satisfies CachedPull });
}

export async function cachedPull(jobNumber: string): Promise<CachedPull | null> {
  const stored = await chrome.storage.local.get(PULL_KEY(jobNumber));
  const pull = stored[PULL_KEY(jobNumber)] as CachedPull | undefined;
  return pull && typeof pull === "object" && pull.csvs ? pull : null;
}

/** A pull already in memory, named the way the folder files are named. */
export function sourcesFromPull(jobNumber: string, csvs: OpenCapCsvs): FolderCsvs {
  const sources: OpenCapSource[] = [];
  for (const spec of OPEN_CAP_FILES) {
    const csv = csvs[spec.kind];
    if (!csv) continue;
    const name = `fieldcap-job-${jobNumber}-${EXPORTER_SUFFIX[spec.kind]}.csv`;
    sources.push({ kind: spec.kind, csv, name, path: `browser-cache\\${name}` });
  }
  return { sources, folderPath: "browser-cache" };
}

export interface SurveyStation {
  md: number;
  inc: number;
  azm: number;
  tvd?: number;
  bit?: number;
}

/** Directional stations for one FieldCap job. Empty when that entity has no depth, inclination and azimuth. */
export async function fieldCapSurveys(clientJobId: string): Promise<SurveyStation[]> {
  if (!globalThis.chrome?.runtime?.id) return [];
  const reply = await chrome.runtime.sendMessage({ type: "phx-fieldcap-surveys", clientJobId }) as { ok?: boolean; stations?: SurveyStation[]; error?: string } | null;
  if (!reply?.ok) throw new Error(reply?.error ?? "FieldCap surveys did not answer.");
  return reply.stations ?? [];
}

/** FieldCap's Jobs list (read-only). Throws with FieldCap's reason when it does not answer. */
export async function fieldCapJobs(): Promise<FieldCapJob[]> {
  if (!globalThis.chrome?.runtime?.id) return [];
  const reply = await chrome.runtime.sendMessage({ type: "phx-fieldcap-jobs" }) as { ok?: boolean; jobs?: FieldCapJob[]; error?: string } | null;
  if (!reply?.ok) throw new Error(reply?.error ?? "FieldCap did not answer.");
  return reply.jobs ?? [];
}
