import type { FieldCapJob } from "./openCapFiles";

export interface JobFolder {
  jobId: string;
  label: string;
}

const PICK_FOLDER = "pick";

const mentions = (label: string, jobNumber: string) => new RegExp(`(^|\\D)${jobNumber}(\\D|$)`).test(label);

/** Field Active / Field Staging, or (when FieldCap gives no status) a job with no end date yet. */
export function isCurrent(job: FieldCapJob): boolean {
  return job.status ? /active|staging/i.test(job.status) : !job.ended;
}

/** The app folder that holds a Job Number: the one it was linked to, else a folder whose name carries the number. */
export function folderFor(jobNumber: string, folders: JobFolder[], links: Record<string, string>): string {
  const linked = links[jobNumber];
  if (linked && folders.some((folder) => folder.jobId === linked)) return linked;
  return folders.find((folder) => mentions(folder.label, jobNumber))?.jobId ?? "";
}

/** The FieldCap job the open folder belongs to. */
export function jobInFolder(appJobId: string, jobs: FieldCapJob[], folders: JobFolder[], links: Record<string, string>): FieldCapJob | null {
  if (!appJobId) return null;
  return jobs.find((job) => folderFor(job.jobNumber, folders, links) === appJobId) ?? null;
}

function statusRank(status: string): number {
  if (/active/i.test(status)) return 0;
  if (/staging/i.test(status)) return 1;
  return 2;
}

function jobLabel(job: FieldCapJob, hasFolder: boolean): string {
  const name = job.name.startsWith(job.jobNumber) ? job.name.slice(job.jobNumber.length).replace(/^\s*[-·]\s*/, "") : job.name;
  return `${job.jobNumber} · ${name || "(no name)"}${job.status ? ` · ${job.status}` : ""}${hasFolder ? "" : " · no folder yet"}`;
}

interface Props {
  /** FieldCap's jobs; null until FieldCap answers. */
  jobs: FieldCapJob[] | null;
  /** Why FieldCap's list is missing, when it is. */
  jobsError: string;
  folders: JobFolder[];
  links: Record<string, string>;
  appJobId: string;
  /** A FieldCap job selected that has no folder yet (Pull and Publish work on it without one). */
  picked: FieldCapJob | null;
  busy: boolean;
  onOpen: (appJobId: string) => void;
  onPick: (job: FieldCapJob) => void;
  /** Ask for a folder: for the picked job, or (null) any other job folder. */
  onChooseFolder: (job: FieldCapJob | null) => void;
}

/** The one job list: FieldCap's Active and Staging jobs first, then any other job that has a folder. */
export function JobPicker({ jobs, jobsError, folders, links, appJobId, picked, busy, onOpen, onPick, onChooseFolder }: Props) {
  const list = jobs ?? [];
  const current = list.filter(isCurrent).sort((a, b) => statusRank(a.status) - statusRank(b.status) || b.jobNumber.localeCompare(a.jobNumber));
  const older = list.filter((job) => !isCurrent(job) && folderFor(job.jobNumber, folders, links) !== "");
  const claimed = new Set([...current, ...older].map((job) => folderFor(job.jobNumber, folders, links)).filter(Boolean));
  const loose = folders.filter((folder) => !claimed.has(folder.jobId));
  const open = picked ?? jobInFolder(appJobId, list, folders, links);
  const value = open ? `job:${open.jobNumber}` : appJobId ? `folder:${appJobId}` : "";

  function choose(next: string) {
    if (next === PICK_FOLDER) {
      onChooseFolder(picked);
      return;
    }
    if (next.startsWith("folder:")) {
      onOpen(next.slice("folder:".length));
      return;
    }
    const job = list.find((candidate) => `job:${candidate.jobNumber}` === next);
    if (!job) return;
    const folder = folderFor(job.jobNumber, folders, links);
    if (folder) onOpen(folder);
    else onPick(job);
  }

  const option = (job: FieldCapJob) => {
    const hasFolder = folderFor(job.jobNumber, folders, links) !== "";
    return <option key={job.clientJobId} value={`job:${job.jobNumber}`}>{jobLabel(job, hasFolder)}</option>;
  };

  return (
    <select className="job-picker" aria-label="Job" value={value} disabled={busy} onChange={(event) => choose(event.target.value)}>
      <option value="" disabled>{jobs === null && !jobsError ? "Loading FieldCap jobs…" : "Choose a job"}</option>
      {jobsError ? <option value="" disabled>FieldCap: {jobsError}</option> : null}
      {current.length > 0 ? <optgroup label="Field Active / Field Staging">{current.map(option)}</optgroup> : null}
      {older.length > 0 ? <optgroup label="Other jobs with a folder">{older.map(option)}</optgroup> : null}
      {loose.length > 0 ? (
        <optgroup label="Folders">
          {loose.map((folder) => <option key={folder.jobId} value={`folder:${folder.jobId}`}>{folder.label}</option>)}
        </optgroup>
      ) : null}
      <option value={PICK_FOLDER}>{picked ? `Choose the folder for ${picked.jobNumber}…` : "Open another job folder…"}</option>
    </select>
  );
}
