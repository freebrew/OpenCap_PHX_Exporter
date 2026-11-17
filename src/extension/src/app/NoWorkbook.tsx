/** Why the grid is empty, in the order the reasons are checked. */
export type NoWorkbookReason =
  | { kind: "no-job" }
  /** A FieldCap job picked in the list that has no folder yet. */
  | { kind: "no-folder"; jobNumber: string }
  /** The import panel above is waiting for Continue. */
  | { kind: "staged"; fileName: string }
  /** The folder has a workbook that was never opened here (it arrived later, or its import was cancelled). */
  | { kind: "unopened"; xlsm: string }
  | { kind: "empty-folder"; label: string };

export function noWorkbookReason(input: { appJobId: string; label: string; pickedJobNumber: string; stagedFile: string; folderXlsm: string }): NoWorkbookReason {
  if (input.stagedFile) return { kind: "staged", fileName: input.stagedFile };
  if (input.pickedJobNumber) return { kind: "no-folder", jobNumber: input.pickedJobNumber };
  if (!input.appJobId) return { kind: "no-job" };
  if (input.folderXlsm) return { kind: "unopened", xlsm: input.folderXlsm };
  return { kind: "empty-folder", label: input.label || "job" };
}

interface Props {
  reason: NoWorkbookReason;
  busy: boolean;
  onOpenWorkbook: (xlsm: string) => void;
  /** Read the open job's folder again. */
  onCheckFolder: () => void;
}

/** The grid has no workbook. Say why, in the grid, and what gets one onto it. */
export function NoWorkbook({ reason, busy, onOpenWorkbook, onCheckFolder }: Props) {
  switch (reason.kind) {
    case "no-job":
      return (
        <section className="no-workbook">
          <h1>No job is open</h1>
          <p>Choose a job in the job list above.</p>
        </section>
      );
    case "no-folder":
      return (
        <section className="no-workbook">
          <h1>Job {reason.jobNumber} has no folder yet</h1>
          <p>In the job list, choose "Choose the folder for {reason.jobNumber}…" and pick the job's folder. Pull from FieldCap and Publish to StrapSheets work without one.</p>
        </section>
      );
    case "staged":
      return (
        <section className="no-workbook">
          <h1>{reason.fileName} is ready to open</h1>
          <p>Press Continue in the panel above.</p>
        </section>
      );
    case "unopened":
      return (
        <section className="no-workbook">
          <h1>{reason.xlsm} has not been opened here</h1>
          <p>It is in this job's folder. Opening it starts a browser shift. The Excel file is not changed.</p>
          <button type="button" disabled={busy} onClick={() => onOpenWorkbook(reason.xlsm)}>Open {reason.xlsm}</button>
        </section>
      );
    case "empty-folder":
      return (
        <section className="no-workbook">
          <h1>No Slide Sheet in the {reason.label} folder</h1>
          <p>Put this job's Slide Sheet workbook (.xlsm) in that folder, then press Check Folder Again.</p>
          <p>Until then, Pull from FieldCap saves its files into the folder's OpenCap folder and loads nothing.</p>
          <button type="button" disabled={busy} onClick={onCheckFolder}>Check Folder Again</button>
        </section>
      );
    default: {
      const never: never = reason;
      throw new Error(`Unhandled empty grid reason ${String(never)}`);
    }
  }
}
