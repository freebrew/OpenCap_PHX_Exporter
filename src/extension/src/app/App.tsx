import { useEffect, useRef, useState } from "react";
import { createEnvelope, directoryFromHandle, grantJobFolder, JobSession, MemoryDirectory, mirrorSnapshot, openJob, openRegistry, parseEnvelope, reopenStoredFolder, sha256, type JobRegistry, type PhxEnvelope, type Recovery, type SaveState, type StoredDirectory } from "@phx/job-store";
import { buildBundleFiles, canonicalize, diffVisibleUserData, isCanonBook, macroButtonsFromWorkbook, mergeVisibleUserData, readWorkbook, restyleWorkbook, STYLE_VERSION, type CanonBook, type ImportLog, type MacroButton, type MappedWorkbook, type VisibleMerge } from "@phx/xlsx-bridge";
import { authorityOf } from "../../../packages/cutover/src/index";
import type { SurveyStation } from "../../../packages/fieldcap-sync/src/index";
import {
  applyDayRoll,
  applyOpenCap,
  backfillDayBreakdown,
  dayBreakdownWarning,
  projectFieldCapDashboard,
  clearTallyInputs,
  computePlanGauge,
  DAY_FIRST,
  DAY_LAST,
  isTrackerSheetName,
  jobField,
  onTrackerEdit,
  applyContDiHeldReadings,
  paintContDiColumnA,
  parseArea,
  readJobFields,
  readTally,
  repairSurveyFormulas,
  SheetBook,
  startDayRoll,
  highlightNewestSurvey,
  refreshSlideComments,
  sortCrewByRole,
  storeTarget,
  surveyBarState,
  syncPlanTargetWindow,
  bindCostHeadline,
  bindLiveToolHours,
  bindSurveyDepthMetres,
  syncLastSurveyStrip,
  toolHoursSync,
  trackerSheet,
  updateWellborePosition,
  writeEntry,
  writeSurveyStations,
  writeSettings,
  type DayRollContext,
  type DayRollInputs,
  type DayWrite,
  type DlEntryStart,
  type GaugeModel,
  type NamedTarget,
  type ProtectionInfo,
  type SnapshotBook,
  type TallyView,
} from "../../../packages/workbook-commands/src/index";
import { createWorkbookHost } from "../workbook/createUniver";
import type { WorkbookHost } from "../workbook/facade";
import { preserveActiveSheet } from "../workbook/preserveActiveSheet";
import { DayRollDialog } from "./DayRollDialog";
import { DlEntryDialog } from "./DlEntryDialog";
import { appJobIdOf, folderWorkbooks, matchingWorkbook, snapshotLabel, staleEntriesFor } from "./jobFolders";
import { JobPicker, jobInFolder } from "./JobPicker";
import { NoWorkbook, noWorkbookReason } from "./NoWorkbook";
import { cachedPull, cachePull, fieldCapJobs, fieldCapSurveys, findFolderCsvs, jobMap, moveJobLinks, pullFromFieldCap, rememberJob, sourcesFromPull, writeFolderCsvs, type FieldCapJob, type FolderCsvs } from "./openCapFiles";
import { crewFromCsv, publishSnapshot, signInToStrapSheets, strapSession } from "./strapSheets";
import type { FolderEntry } from "./pickFile";
import { PipeTallyDialog } from "./PipeTallyDialog";
import { SheetOverlay } from "./SheetOverlay";
import { appendTally, contDiEdited, emailReport, importTallyCsv, runMacro, type MacroContext, type ToolSection } from "./sheetMacros";
import { TargetsDialog, type TargetEdit } from "./TargetsDialog";
import { Tools, TOOLS_DOCK_ID } from "./Tools";
import { WellsView } from "./WellsView";
import { commitHandoff, latestReceipt, pendingHandoffs, workbookToken } from "./handoffFolder";
import { SyncConflictDialog } from "./SyncConflictDialog";
import { opWrites, scalarWrites } from "./syncWrites";
import { PREVIEW_VERSION } from "../version";

interface StagedImport {
  mapped: MappedWorkbook;
  fileName: string;
}

interface SyncReview {
  fileName: string;
  jobLabel: string;
  hash: string;
  excel: CanonBook;
  merge: VisibleMerge;
  preferBrowserMachine: boolean;
}

interface FolderHandle {
  getFileHandle(name: string, options?: { create?: boolean }): Promise<{
    getFile(): Promise<Blob>;
    createWritable(options: { mode?: "exclusive" | "siloed" }): Promise<{
      write(data: BufferSource | Blob | string): Promise<void>;
      close(): Promise<void>;
      abort(): Promise<void>;
    }>;
  }>;
  getDirectoryHandle(name: string, options?: { create?: boolean }): Promise<FolderHandle>;
  entries(): AsyncIterable<[string, { kind: string }]>;
  queryPermission(descriptor: { mode: "readwrite" }): Promise<"granted" | "prompt" | "denied">;
  requestPermission(descriptor: { mode: "readwrite" }): Promise<"granted" | "prompt" | "denied">;
}

function pickedDirectoryHandle(): Promise<FolderHandle> {
  const picker = (window as Window & { showDirectoryPicker?: (options: { mode: "readwrite"; startIn: "desktop" }) => Promise<FolderHandle> }).showDirectoryPicker;
  if (!picker) return Promise.reject(new Error("This browser cannot grant a job folder."));
  return picker({ mode: "readwrite", startIn: "desktop" });
}

async function adapterDirectory(): Promise<StoredDirectory | null> {
  if (!import.meta.env.DEV) return null;
  const adapter = new URLSearchParams(window.location.search).get("adapter");
  if (adapter === "memory") return new MemoryDirectory();
  // OPFS handles are always granted, so they carry no permission methods.
  if (adapter === "opfs") return directoryFromHandle(await navigator.storage.getDirectory() as unknown as Parameters<typeof directoryFromHandle>[0]);
  return null;
}

/** The OPFS test folder as a job folder, so folder macros (REFRESH CSVs) run under test. */
/** Chrome's NotFoundError: the granted folder was moved, renamed, or is on a drive that is not connected. */
function folderGone(error: unknown): boolean {
  return error instanceof DOMException && error.name === "NotFoundError";
}

function missingFolderText(label: string): string {
  return `The job folder for "${label}" cannot be found. It may be on a drive that is not connected (a USB or external drive), or it was moved or renamed.\n\nReconnect the drive and try again, or choose "Open another job folder…" in the job list to pick the folder where it is now.`;
}

/** How often the job list asks FieldCap again, so a job that goes to Staging or Active shows up. */
const JOBS_REFRESH_MS = 120_000;
/** How often the open job asks FieldCap for new surveys. One small query; the sheet is written only when the stations changed. */
const SURVEY_POLL_MS = 60_000;

async function adapterJobFolder(): Promise<FolderEntry | null> {
  if (!import.meta.env.DEV || new URLSearchParams(window.location.search).get("adapter") !== "opfs") return null;
  return await navigator.storage.getDirectory() as unknown as FolderEntry;
}

export function App() {
  const grid = useRef<HTMLDivElement>(null);
  const host = useRef<WorkbookHost | null>(null);
  const session = useRef<JobSession | null>(null);
  const directory = useRef<StoredDirectory | null>(null);
  const [jobs, setJobs] = useState<Array<{ jobId: string; label: string }>>([]);
  const [fieldCapList, setFieldCapList] = useState<FieldCapJob[] | null>(null);
  const [fieldCapError, setFieldCapError] = useState("");
  const [links, setLinks] = useState<Record<string, string>>({});
  const [pickedJob, setPickedJob] = useState<FieldCapJob | null>(null);
  const adapterFolder = useRef<StoredDirectory | null>(null);
  const [status, setStatus] = useState<SaveState>("folder-safe");
  const [message, setMessage] = useState("Choose a job in the job list to open this migration preview.");
  const [recovery, setRecovery] = useState<Recovery>({ kind: "none" });
  const [staged, setStaged] = useState<StagedImport | null>(null);
  const [backups, setBackups] = useState(0);
  const [jobId, setJobId] = useState("");
  const [authority, setAuthority] = useState<"excel" | "native">("excel");
  const [buttons, setButtons] = useState<MacroButton[]>([]);
  const [sheetName, setSheetName] = useState("");
  const [surface, setSurface] = useState<"sheets" | "wells">("sheets");
  const [toolRequest, setToolRequest] = useState<{ section: ToolSection; token: number } | null>(null);
  const [dayRoll, setDayRoll] = useState<DayRollContext | null>(null);
  const dayRollQueue = useRef<number[]>([]);
  const dayRollOpen = useRef(false);
  const syncing = useRef(false);
  const refreshDashboardRef = useRef<() => Promise<void>>(async () => undefined);
  const pollSurveysRef = useRef<() => Promise<void>>(async () => undefined);
  /** The FieldCap stations the poll last put on the sheet. Unchanged stations are not written again over the user's edits. */
  const surveyPollKey = useRef("");
  const [workbookHost, setWorkbookHost] = useState<WorkbookHost | null>(null);
  const [gauge, setGauge] = useState<GaugeModel | null>(null);
  const [busy, setBusy] = useState(false);
  const [menusShown, setMenusShown] = useState(false);
  const [dlEntry, setDlEntry] = useState<{ start: DlEntryStart; book: SheetBook } | null>(null);
  const [targets, setTargets] = useState<NamedTarget[] | null>(null);
  const [pipe, setPipe] = useState<{ view: TallyView; status: string } | null>(null);
  const [ocJobId, setOcJobId] = useState("");
  const handoffBlocks = useRef<string[]>([]);
  const [shiftReady, setShiftReady] = useState(false);
  const [handoffNote, setHandoffNote] = useState("");
  const [pendingHandoff, setPendingHandoff] = useState("");
  const [driftFile, setDriftFile] = useState("");
  const [syncReview, setSyncReview] = useState<SyncReview | null>(null);
  const dismissedDrift = useRef("");
  const inboundLock = useRef(false);
  const snapshotFileRef = useRef("");
  /** True after a FieldCap pull or survey poll wrote this shift. Machine-written conflicts then default to the browser. */
  const machinePreferBrowser = useRef(false);
  const refreshHandoffRef = useRef<(folder: StoredDirectory, force?: boolean) => Promise<boolean>>(async () => false);
  /** A workbook is on the grid. When none is, NoWorkbook says why in its place. */
  const [bookOpen, setBookOpen] = useState(false);
  /** The .xlsm in the open folder that has no snapshot yet. */
  const [folderXlsm, setFolderXlsm] = useState("");

  function stamp(workbook: unknown, sidecar: unknown, patch?: { hash?: string; baseline?: unknown; blocks?: string[] }): PhxEnvelope {
    const previous = session.current?.current;
    const blocks = patch?.blocks ?? handoffBlocks.current;
    return createEnvelope({
      univerAppVersion: previous?.univerAppVersion || "1.3.1",
      sourceWorkbookHash: patch?.hash ?? previous?.sourceWorkbookHash ?? "",
      workbook,
      sidecar,
      importBaseline: patch && "baseline" in patch ? patch.baseline : previous?.importBaseline,
      handoffBlocks: blocks,
    });
  }

  useEffect(() => {
    let cancelled = false;
    void adapterDirectory().then((folder) => {
      if (!cancelled && folder) {
        directory.current = folder;
        adapterFolder.current = folder;
      }
    });
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    let stopped = false;
    const load = async () => {
      const [list, map] = await Promise.all([
        fieldCapJobs().then((value) => ({ value, error: "" }), (error: unknown) => ({ value: null, error: error instanceof Error ? error.message : String(error) })),
        jobMap().catch(() => ({})),
      ]);
      if (stopped) return;
      setLinks(map);
      setFieldCapError(list.error);
      if (list.value) setFieldCapList(list.value);
    };
    void load();
    const timer = window.setInterval(() => void load(), JOBS_REFRESH_MS);
    return () => {
      stopped = true;
      window.clearInterval(timer);
    };
  }, []);

  useEffect(() => {
    const poll = () => void pollSurveysRef.current();
    const timer = window.setInterval(poll, SURVEY_POLL_MS);
    window.addEventListener("focus", poll);
    return () => {
      window.clearInterval(timer);
      window.removeEventListener("focus", poll);
    };
  }, []);

  useEffect(() => {
    const onFocus = () => {
      const folder = directory.current;
      if (folder) void refreshHandoffRef.current(folder, false);
    };
    window.addEventListener("focus", onFocus);
    return () => window.removeEventListener("focus", onFocus);
  }, []);

  useEffect(() => {
    let disposed = false;
    const node = grid.current;
    if (!node) return undefined;
    const created = createWorkbookHost(node);
    host.current = created;
    setWorkbookHost(created);
    if (import.meta.env.DEV) {
      const bridge = window as Window & {
        __phx?: {
          directory: StoredDirectory | null;
          setCell: WorkbookHost["setCell"];
          getCell: WorkbookHost["getCell"];
          getRawCell: WorkbookHost["getRawCell"];
          flushCalculation: WorkbookHost["flushCalculation"];
          snapshot: WorkbookHost["snapshot"];
        };
      };
      bridge.__phx = {
        get directory() {
          return directory.current;
        },
        setCell: (sheetName, a1, value) => created.setCell(sheetName, a1, value),
        getCell: (sheetName, a1) => created.getCell(sheetName, a1),
        getRawCell: (sheetName, a1) => created.getRawCell(sheetName, a1),
        flushCalculation: (timeoutMs?: number) => created.flushCalculation(timeoutMs),
        snapshot: () => created.snapshot(),
      };
    }
    setSheetName(created.activeSheet());
    const stopSheet = created.onActiveSheet(setSheetName);
    // Sheet48: the newest-survey bar, the slide comments and the gauge refresh when the
    // selection leaves the Inc/Azm cells (F13:G505), not on every keystroke inside them.
    let wasInSurvey = false;
    const inSurveyCells = (range: { startRow: number; endRow: number; startColumn: number; endColumn: number }) =>
      range.startColumn <= 6 && range.endColumn >= 5 && range.endRow >= 12 && range.startRow <= 504;
    const stopSelection = created.onSelectionChanged((editedSheet, range) => {
      const nowIn = editedSheet === "Slidesheet" && inSurveyCells(range);
      if (wasInSurvey && !nowIn) void macroContext(created).redrawAfterSurvey();
      wasInSurvey = nowIn;
    });
    const stopEdits = created.onCellsEdited((editedSheet, range) => {
      // Writes wait until Univer has finished the edit that raised this event.
      if (isTrackerSheetName(editedSheet)) {
        window.setTimeout(() => trackerEdited(editedSheet, range), 0);
        return;
      }
      if (editedSheet === "Cont DI") {
        window.setTimeout(() => contDiTargetEdited(range), 0);
        return;
      }
      // SortCrewByRole: a change in the crew manifest's role column reorders the rows.
      const crew = namedCell("OC_CrewTop");
      const crewBottom = namedCell("OC_CrewBot");
      if (editedSheet === "Setup" && crew && crewBottom && range.startColumn <= 9 && range.endColumn >= 9 && range.endRow + 1 >= crew.row && range.startRow + 1 <= crewBottom.row) {
        window.setTimeout(() => resortCrew(crew.row, crewBottom.row), 0);
      }
      if (editedSheet === "Data" && range.startRow <= 2 && range.endRow >= 2 && range.startColumn <= 7 && range.endColumn >= 7) {
        window.setTimeout(() => void refreshDashboardRef.current(), 0);
      }
      const periodColumn = 7;
      if (editedSheet !== "Data" || range.startColumn > periodColumn || range.endColumn < periodColumn) return;
      for (let row = Math.max(range.startRow + 1, DAY_FIRST); row <= Math.min(range.endRow + 1, DAY_LAST); row += 1) dayRollQueue.current.push(row);
      window.setTimeout(nextDayRoll, 0);
    });
    // The green bar and the active target follow the last survey row painted the survey green. Fill
    // Color, Undo and a paste raise no edit event, so every settled change looks at where the bar
    // belongs. A redraw that finds the bar already right does nothing, so its own writes do not loop.
    let barSeenAt = -1;
    let barBusy = false;
    let barTriedAt = -1;
    const followSurveyBar = () => {
      if (barBusy) return;
      const book = freshBook();
      if (!book || !book.hasSheet("Slidesheet")) return;
      const { newest, stale } = surveyBarState(book);
      const moved = barSeenAt >= 0 && newest !== barSeenAt;
      barSeenAt = newest;
      if (!stale) barTriedAt = -1;
      // One try per row: a bar the sheet will not take must not redraw every 300 ms for ever.
      if ((!stale && !moved) || (stale && !moved && barTriedAt === newest)) return;
      barTriedAt = stale ? newest : -1;
      barBusy = true;
      macroContext(created).redrawAfterSurvey().catch((error: unknown) => console.error("Survey bar", error)).finally(() => {
        barBusy = false;
      });
    };
    const stop = created.onDirty(() => {
      const current = session.current;
      const snap = created.snapshot();
      followSurveyBar();
      if (snap) refreshGauge(snap);
      if (!current || !snap) return;
      const previous = current.current;
      current.noteEdit(stamp(snap, previous?.sidecar ?? {}));
      setStatus(current.state);
    });
    const stopBlocked = created.onBlockedCommand((id, executed) => {
      if (executed && !handoffBlocks.current.includes(id)) handoffBlocks.current = [...handoffBlocks.current, id];
      setHandoffNote(executed
        ? `${id} changed the sheet in a way Excel cannot take back. Export is blocked until you re-import.`
        : `${id} was refused. Excel cannot take that kind of change back.`);
    });
    openRegistry().then(async (registry) => {
      if (disposed) return;
      setJobs(await registry.list());
    }).catch((error: unknown) => setMessage(error instanceof Error ? error.message : String(error)));
    return () => {
      disposed = true;
      stop();
      stopBlocked();
      stopSheet();
      stopSelection();
      stopEdits();
      created.dispose();
      session.current?.dispose();
    };
  }, []);

  useEffect(() => {
    const onKey = (event: KeyboardEvent) => {
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "s") {
        event.preventDefault();
        void saveNow();
      }
    };
    const onLeave = (event: BeforeUnloadEvent) => {
      if (session.current?.dirty() || session.current?.folderUnsafe()) {
        event.preventDefault();
        event.returnValue = "";
      }
    };
    window.addEventListener("keydown", onKey);
    window.addEventListener("beforeunload", onLeave);
    return () => {
      window.removeEventListener("keydown", onKey);
      window.removeEventListener("beforeunload", onLeave);
    };
  }, []);

  function writeData(writes: DayWrite[]) {
    for (const write of writes) {
      if ("formula" in write) {
        const area = parseArea(write.a1);
        if (area) host.current?.applyWrites([{ kind: "formula", sheet: "Data", row: area.row, col: area.col, formula: write.formula }]);
      } else if ("format" in write) {
        const area = parseArea(write.a1);
        if (area) host.current?.applyWrites([{ kind: "format", sheet: "Data", row: area.row, col: area.col, pattern: write.format }]);
      } else {
        host.current?.setCell("Data", write.a1, write.value);
      }
    }
  }

  /** RefreshData's closing check: the day rows against FieldCap's days, after they recalculate. */
  async function checkDayBreakdown(): Promise<string> {
    const created = host.current;
    if (!created) return "";
    await created.flushCalculation();
    const snap = created.snapshot() as SnapshotBook | null;
    return snap ? dayBreakdownWarning(snap) : "";
  }

  /** Locked / unlocked cells as the import recorded them; macros skip locked cells like the VBA. */
  function protection(): ProtectionInfo[] {
    const sidecar = session.current?.current?.sidecar;
    const list = sidecar && typeof sidecar === "object" ? (sidecar as { protection?: unknown }).protection : null;
    return Array.isArray(list) ? list as ProtectionInfo[] : [];
  }

  /** A defined name from the import ("OC_ImpPlanRow" -> Setup!$J$29) as a 1-based cell. */
  function namedCell(name: string): { row: number; col: number } | null {
    const sidecar = session.current?.current?.sidecar;
    const names = sidecar && typeof sidecar === "object" ? (sidecar as { names?: unknown }).names : null;
    if (!Array.isArray(names)) return null;
    const record = (names as Array<{ name?: string; text?: string }>).find((entry) => entry.name?.toLowerCase() === name.toLowerCase());
    const area = record?.text ? parseArea(record.text.replace(/^=/, "").replace(/^'?[^!]*'?!/, "")) : null;
    return area ? { row: area.row, col: area.col } : null;
  }

  async function jobFolder(appJobId = jobId): Promise<FolderEntry | null> {
    if (!appJobId) return null;
    const handle = (await openRegistry().then((registry) => registry.get(appJobId)).catch(() => null))?.handle as Partial<FolderEntry> | undefined;
    return handle?.kind === "directory" && typeof handle.resolve === "function" ? handle as FolderEntry : adapterJobFolder();
  }

  function freshBook(): SheetBook | null {
    const snap = host.current?.snapshot() as SnapshotBook | null;
    return snap ? new SheetBook(snap, protection()) : null;
  }

  /**
   * A workbook that arrived from Excel can hold survey formulas a deleted cell broke (#REF!, or a
   * row reading the Inc one row too high). The cumulative N/E columns carry that error down every
   * row, and the plan gauge and the slide advice start from it. Rewrite them from the column's
   * own pattern, recalculate and redraw. Returns the sentence for the status line, or "".
   */
  async function repairDamagedSurveyFormulas(): Promise<string> {
    const created = host.current;
    const book = freshBook();
    if (!created || !book) return "";
    const painted = paintContDiColumnA(book);
    const held = applyContDiHeldReadings(book);
    const circ = bindLiveToolHours(book);
    const strip = syncLastSurveyStrip(book);
    const depth = bindSurveyDepthMetres(book);
    const costs = bindCostHeadline(book);
    const fixed = repairSurveyFormulas(book);
    if (!painted && !held && !circ && !strip && !depth && !costs && fixed.length === 0) return "";
    created.applyWrites(book.writes);
    if (fixed.length === 0) {
      if (held || circ || strip || depth || costs) await created.flushCalculation();
      return "";
    }
    await macroContext(created).redrawAfterSurvey();
    const rows = [...new Set(fixed.map((cell) => cell.row))].sort((a, b) => a - b);
    return ` Repaired ${fixed.length} damaged survey formula${fixed.length === 1 ? "" : "s"} on Slidesheet row${rows.length === 1 ? "" : "s"} ${rows.slice(0, 12).join(", ")}${rows.length > 12 ? ", …" : ""} (a cell deleted in Excel).`;
  }

  function sidecarRecord(): Record<string, unknown> {
    const sidecar = session.current?.current?.sidecar;
    return sidecar && typeof sidecar === "object" ? sidecar as Record<string, unknown> : {};
  }

  /** GetNamedText: a name that holds text (="21784") rather than a cell. */
  function namedText(name: string): string {
    const names = sidecarRecord().names;
    const record = Array.isArray(names) ? (names as Array<{ name?: string; text?: string }>).find((entry) => entry.name?.toLowerCase() === name.toLowerCase()) : undefined;
    const text = (record?.text ?? "").replace(/^=/, "");
    return text.length >= 2 && text.startsWith("\"") && text.endsWith("\"") ? text.slice(1, -1) : text.replace(/"/g, "");
  }

  /** SetNamedText: a blank value deletes the name. */
  function setNamedText(name: string, value: string) {
    const current = session.current;
    const previous = current?.current;
    const snap = host.current?.snapshot();
    if (!current || !previous || !snap) return;
    const sidecar = sidecarRecord();
    const names = Array.isArray(sidecar.names) ? sidecar.names as Array<Record<string, unknown>> : [];
    const safe = value.replace(/^ +| +$/g, "").replace(/"/g, "");
    const old = names.find((entry) => String(entry.name ?? "").toLowerCase() === name.toLowerCase());
    const kept = names.filter((entry) => entry !== old);
    const record = { hidden: false, builtin: false, supported: true, ...old, name, text: `"${safe}"` };
    current.noteEdit(stamp(snap, { ...sidecar, names: safe ? [...kept, record] : kept }));
    setStatus(current.state);
  }

  /** The Job Number the workbook holds ("Job Code", then "Job ID"). What the toolbar compares. */
  function workbookOcJobId(): string {
    const book = freshBook();
    return book ? jobField(readJobFields(book), "Job Code", "Job ID").trim() : "";
  }

  /** RefreshAllCsvData on the CSVs found, then the BHA mirror, motor, hours tables and day rows. */
  async function loadOpenCap(found: FolderCsvs): Promise<string | null> {
    const created = host.current;
    const book = freshBook();
    if (!created || !book) return null;
    const outer = syncing.current;
    syncing.current = true;
    try {
      const result = applyOpenCap(book, found.sources, {
        lastJobId: namedText("OC_LastJobId"),
        folderPath: found.folderPath,
        crewTop: namedCell("OC_CrewTop")?.row,
        crewBottom: namedCell("OC_CrewBot")?.row,
      });
      // Advance to the newest/active BHA from FieldCap, but never downgrade a higher BHA already on the same job.
      const dash = projectFieldCapDashboard(book, { noDowngrade: !result.newJob });
      created.applyWrites(book.writes);
      await created.flushCalculation();
      const snap = created.snapshot() as SnapshotBook | null;
      if (snap) writeData(backfillDayBreakdown(snap));
      const dayWarning = await checkDayBreakdown();
      if (result.jobId) setNamedText("OC_LastJobId", result.jobId);
      setOcJobId(result.jobId);
      refreshGauge(created.snapshot());
      const bhaNote = dash && dash.bhas > 0 ? ` BHA ${dash.activeBha || "—"} on H3, motor ${dash.motor || "—"} on C22. Day rows follow that BHA.` : "";
      setMessage(`${result.status}${bhaNote}`);
      if (result.alert) window.alert(result.alert);
      if (dayWarning) window.alert(dayWarning);
      return result.jobId;
    } finally {
      syncing.current = outer;
    }
  }

  /**
   * quiet = the background poll: it never asks or alerts, and says nothing when nothing changed.
   * Returns false when the stations were refused, so the poll tries them again.
   */
  async function applyFieldCapSurveys(jobNumber: string, stations: readonly SurveyStation[], quiet = false): Promise<boolean> {
    const book = freshBook();
    if (!book) return false;
    if (stations.length === 0) {
      if (!quiet) setMessage(`Job ${jobNumber}: FieldCap has no survey stations for this job.`);
      return true;
    }
    let result = writeSurveyStations(book, stations);
    if (result.needsConfirmation && quiet) {
      setMessage(`Job ${jobNumber}: FieldCap has ${result.tvdMismatches.length} past-survey correction(s) waiting. Press Pull from FieldCap to review them.`);
      return true;
    }
    if (result.needsConfirmation) {
      const details = result.tvdMismatches.slice(0, 5)
        .map((entry) => `Row ${entry.row}, MD ${entry.md}: ${entry.difference.toFixed(2)} m`)
        .join("\n");
      const more = result.tvdMismatches.length > 5 ? `\n…and ${result.tvdMismatches.length - 5} more` : "";
      const confirmed = window.confirm(
        `FieldCap would correct ${result.tvdMismatches.length} past survey(s) (their direction or their depth) whose TVD differs from the Slidesheet by more than 0.10 m.\n\n${details}${more}\n\nFieldCap is the authority. Continue?`,
      );
      if (!confirmed) {
        setMessage(`Job ${jobNumber}: past-survey corrections cancelled; no surveys changed.`);
        return true;
      }
      result = writeSurveyStations(book, stations, { allowTvdMismatch: true });
    }
    if (result.warning) {
      if (!quiet) window.alert(result.warning);
      setMessage(`Job ${jobNumber}: surveys not changed. ${result.warning}`);
      return false;
    }
    if (book.writes.length > 0) {
      host.current?.applyWrites(book.writes);
      machinePreferBrowser.current = true;
    }
    // Slide the T2:Y5 window and light the active target for the last survey, as when a survey is
    // typed by hand (Sheet48's redraw). It runs even when this pull wrote nothing, so a sheet whose
    // surveys are already in catches up too.
    if (host.current) await macroContext(host.current).redrawAfterSurvey();
    const skipped = result.skippedPlaceholders > 0 ? ` ${result.skippedPlaceholders} untaken trailing survey(s) omitted.` : "";
    const note = result.note ? ` ${result.note}` : "";
    const fixes = result.depthCorrections.length;
    const depths = fixes === 0
      ? ""
      : ` FieldCap's depth replaced ${fixes} Slidesheet bit depth${fixes === 1 ? "" : "s"} (${result.depthCorrections.slice(0, 3).map((fix) => `row ${fix.row}: ${fix.from} to ${fix.to}`).join("; ")}${fixes > 3 ? "; …" : ""}).`;
    const offsets = result.offsetChanges.length === 0
      ? ""
      : ` Survey Offset ${result.offsetChanges.slice(0, 3).map((change) => `${change.from == null ? "set" : `${change.from} to`} ${change.to} at row ${change.row} (MD ${change.md})`).join("; ")}${result.offsetChanges.length > 3 ? "; …" : ""}, from FieldCap's bit depths.`;
    const newest = stations[stations.length - 1];
    const upTo = newest ? ` FieldCap's newest survey is MD ${newest.md}.` : "";
    const changedSheet = result.changed > 0 || result.offsetChanges.length > 0;
    if (!quiet || changedSheet) {
      setMessage(`Job ${jobNumber}: ${result.matched} surveys matched, ${result.appended} appended, ${result.changed} changed.${depths}${offsets}${skipped}${note}${upTo} BHA totals follow Data H3.`);
    }
    if (!quiet && result.note && result.checkOffset) window.alert(result.note);
    return true;
  }

  /** Surveys as they come in: the newest FieldCap stations reach the sheet without a Pull. */
  async function pollFieldCapSurveys(): Promise<void> {
    if (!host.current || !bookOpen || busy || syncing.current || document.hidden) return;
    const held = workbookOcJobId();
    const job = held === "" ? undefined : fieldCapList?.find((entry) => entry.jobNumber === held);
    if (!job) return;
    syncing.current = true;
    try {
      const stations = await fieldCapSurveys(job.clientJobId);
      const key = `${held}|${JSON.stringify(stations)}`;
      if (key === surveyPollKey.current) return;
      if (await applyFieldCapSurveys(held, stations, true)) surveyPollKey.current = key;
    } catch {
      // FieldCap did not answer (signed out, offline): the next poll asks again.
    } finally {
      syncing.current = false;
    }
  }
  pollSurveysRef.current = pollFieldCapSurveys;

  /** H3 changed: the hours tables, C22 and the day rows follow the BHA now in the mirror. */
  async function refreshDashboard(): Promise<void> {
    if (syncing.current) return;
    const created = host.current;
    const book = freshBook();
    if (!created || !book) return;
    syncing.current = true;
    try {
      const dash = projectFieldCapDashboard(book, { preserveSelection: true });
      if (!dash || dash.bhas === 0) return;
      created.applyWrites(book.writes);
      await created.flushCalculation();
      const snap = created.snapshot() as SnapshotBook | null;
      if (snap) writeData(backfillDayBreakdown(snap));
      setMessage(`BHA ${dash.activeBha} is selected. Motor ${dash.motor || "—"} is on Data C22, and the day rows use that BHA.`);
    } finally {
      syncing.current = false;
    }
  }
  refreshDashboardRef.current = refreshDashboard;

  /** Opening a job applies the cached FieldCap pull, then refreshes it. Surveys land on the Slidesheet. */
  async function autosyncOpenJob(appJobId: string): Promise<void> {
    if (syncing.current || !host.current) return;
    const held = workbookOcJobId();
    const links = await jobMap().catch(() => ({}));
    const linked = Object.entries(links).find(([, id]) => id === appJobId)?.[0] ?? "";
    const jobNumber = held || linked;
    if (jobNumber === "" || (held !== "" && linked !== "" && held !== linked)) return;
    syncing.current = true;
    try {
      const cached = await cachedPull(jobNumber);
      if (cached && Object.keys(cached.csvs).length > 0) {
        setMessage(`Job ${jobNumber}: applying the saved FieldCap pull.`);
        syncing.current = false;
        await loadOpenCap(sourcesFromPull(jobNumber, cached.csvs));
        syncing.current = true;
      }
      let jobs: FieldCapJob[] = [];
      try {
        jobs = await fieldCapJobs();
      } catch (error) {
        if (!cached) setMessage(error instanceof Error ? error.message : String(error));
        return;
      }
      const job = jobs.find((entry) => entry.jobNumber === jobNumber);
      if (!job) return;
      setMessage(`Job ${jobNumber}: pulling FieldCap.`);
      const csvs = await pullFromFieldCap(job.clientJobId, (progress) => setMessage(`FieldCap job ${jobNumber}: ${progress.label} ${progress.pct}%`));
      if (Object.keys(csvs).length === 0) {
        setMessage(`FieldCap returned no files for job ${jobNumber}. Open FieldCap in a browser tab and sign in.`);
        return;
      }
      await cachePull(jobNumber, csvs);
      const root = await jobFolder();
      syncing.current = false;
      if (root) await loadOpenCap(await writeFolderCsvs(root, jobNumber, csvs));
      else await loadOpenCap(sourcesFromPull(jobNumber, csvs));
      syncing.current = true;
      const stations = await fieldCapSurveys(job.clientJobId);
      await applyFieldCapSurveys(jobNumber, stations);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : String(error));
    } finally {
      syncing.current = false;
    }
  }

  /** Setup REFRESH CSVs: the newest OpenCap exports in the job folder. */
  async function refreshOpenCap(): Promise<void> {
    const root = await jobFolder();
    if (!root) {
      window.alert("Choose the job in the job list first. Refresh scans <job folder>\\OpenCap\\ for CSV exports.");
      return;
    }
    await loadOpenCap(await findFolderCsvs(root));
  }

  /** The FieldCap job the open folder belongs to (the one the job list shows as selected). */
  const selectedJob = pickedJob ?? jobInFolder(jobId, fieldCapList ?? [], jobs, links);

  /** The OpenCap exporter's pull for the selected job, saved to <job>\OpenCap\ and loaded like REFRESH CSVs. */
  async function pullFieldCap() {
    const pullHost = host.current;
    if (!pullHost || busy || !selectedJob) return;
    const { clientJobId, jobNumber } = selectedJob;
    // A job with no folder yet is pulled into the browser only; the workbook on screen belongs to another job.
    const root = pickedJob ? null : await jobFolder();
    const held = root ? workbookOcJobId() : "";
    if (held && held !== jobNumber && !window.confirm(`This workbook holds FieldCap job ${held}.\n\nLoad job ${jobNumber} into it? Setup's job, crew and files panels, Costs and the mud motors switch to ${jobNumber}.`)) return;
    setBusy(true);
    try {
      await preserveActiveSheet(pullHost, async () => {
        const csvs = await pullFromFieldCap(clientJobId, (progress) => setMessage(`FieldCap job ${jobNumber}: ${progress.label} ${progress.pct}%`));
        if (Object.keys(csvs).length === 0) {
          window.alert(`FieldCap returned no data for job ${jobNumber}. Open FieldCap in a browser tab and sign in, then pull again.`);
          return;
        }
        await cachePull(jobNumber, csvs);
        if (!root) {
          setMessage(`Pulled job ${jobNumber} from FieldCap. Publish to StrapSheets works now. When the job folder exists, choose "Choose the folder for ${jobNumber}…" in the job list and the pull is saved into it.`);
          return;
        }
        if (!freshBook()) {
          // The folder has no workbook on the grid: the pull is kept, and the user is told it went nowhere.
          const saved = await writeFolderCsvs(root, jobNumber, csvs);
          await linkJob(jobNumber, jobId);
          const text = `Job ${jobNumber} was pulled from FieldCap and its ${saved.sources.length} files were saved into ${saved.folderPath}. Nothing was loaded, because no Slide Sheet is open for this job. Open the job's Slide Sheet, then pull again.`;
          setMessage(text);
          window.alert(text);
          return;
        }
        await loadOpenCap(await writeFolderCsvs(root, jobNumber, csvs));
        machinePreferBrowser.current = true;
        await linkJob(jobNumber, jobId);
        const stations = await fieldCapSurveys(clientJobId);
        await applyFieldCapSurveys(jobNumber, stations);
      });
    } catch (error) {
      window.alert(folderGone(error) ? missingFolderText(jobs.find((job) => job.jobId === jobId)?.label ?? "this job") : `FieldCap pull failed: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      setBusy(false);
    }
  }

  async function linkJob(jobNumber: string, appJobId: string) {
    await rememberJob(jobNumber, appJobId);
    setLinks(await jobMap());
  }

  /** Link a folder to a job (asked once; the job list then opens it). A pull made before the folder existed is saved into it. */
  async function chooseFolder(job: FieldCapJob | null) {
    const appJobId = await addFolder();
    if (!appJobId || !job) return;
    await linkJob(job.jobNumber, appJobId);
    const pull = await cachedPull(job.jobNumber);
    const root = pull ? await jobFolder(appJobId) : null;
    if (pull && root) {
      const { sources, folderPath } = await writeFolderCsvs(root, job.jobNumber, pull.csvs);
      setMessage(`Job ${job.jobNumber} is linked to this folder. Its FieldCap pull (${sources.length} files) was saved into ${folderPath}. Setup's REFRESH CSVs or Pull from FieldCap loads it into the workbook.`);
      return;
    }
    setMessage(freshBook()
      ? `Job ${job.jobNumber} is linked to this folder. Press Pull from FieldCap to load its job details, crew and inventory.`
      : `Job ${job.jobNumber} is linked to this folder. It has no Slide Sheet open yet, so Pull from FieldCap would only save its files into the folder.`);
  }

  /** Send the selected job's pulled inventory and crew to StrapSheets: its folder's newest OpenCap CSVs, else the browser's copy of the last pull. FieldCap is not contacted. */
  async function publishStrapSheets() {
    const jobNumber = selectedJob?.jobNumber || ocJobId;
    if (busy || jobNumber === "") return;
    setBusy(true);
    try {
      const root = pickedJob ? null : await jobFolder();
      const sources = root ? (await findFolderCsvs(root)).sources : [];
      const pulled = await cachedPull(jobNumber);
      const inventoryCsv = sources.find((source) => source.kind === "inventory")?.csv ?? pulled?.csvs.inventory ?? "";
      const crewCsv = sources.find((source) => source.kind === "crew")?.csv ?? pulled?.csvs.crew ?? "";
      if (inventoryCsv === "" && crewCsv === "") {
        window.alert(`No inventory or crew has been pulled for job ${jobNumber}. Press Pull from FieldCap first.`);
        return;
      }
      const session = (await strapSession()) ?? (await signInToStrapSheets("microsoft"));
      const result = await publishSnapshot(session, jobNumber, inventoryCsv, crewFromCsv(crewCsv));
      setMessage(`Published job ${result.jobNumber} to StrapSheets: ${result.items} tools (${result.filteredOut} left out by the Data tab filter), ${result.crew} crew.`);
    } catch (error) {
      window.alert(folderGone(error) ? missingFolderText(jobs.find((job) => job.jobId === jobId)?.label ?? `job ${jobNumber}`) : `StrapSheets publish failed: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      setBusy(false);
    }
  }

  /** MDL_PlanGauge.RenderPlanGauge + UpdateWellborePosition on the current values. */
  function refreshGauge(snap: unknown) {
    try {
      const book = new SheetBook(snap as SnapshotBook, protection());
      if (!book.hasSheet("Slidesheet")) {
        setGauge(null);
        return;
      }
      const result = computePlanGauge(book);
      setGauge(result.model);
      if (result.wellbore) updateWellborePosition(book, result.wellbore);
      syncLastSurveyStrip(book);
      bindCostHeadline(book);
      if (book.writes.length > 0) host.current?.applyWrites(book.writes);
    } catch (error) {
      console.error("Plan gauge", error);
    }
  }

  /** DL Tracker Worksheet_Change: page, command or pass/fail typed in the log. */
  function trackerEdited(sheet: string, range: { startRow: number; endRow: number; startColumn: number; endColumn: number }) {
    const book = freshBook();
    if (!book || trackerSheet(book) !== sheet) return;
    onTrackerEdit(book, sheet, { row: range.startRow + 1, col: range.startColumn + 1, row2: range.endRow + 1, col2: range.endColumn + 1 });
    if (book.writes.length > 0) host.current?.applyWrites(book.writes);
  }

  /** SortCrewByRole on the rows the crew manifest spans. */
  function resortCrew(top: number, bottom: number): void {
    const book = freshBook();
    if (!book || !sortCrewByRole(book, top, bottom)) return;
    host.current?.applyWrites(book.writes);
  }

  /** ContDI_ApplyTargetSelection: an edit in the yellow target block rewrites the aim formulas. */
  function contDiTargetEdited(range: { startRow: number; endRow: number; startColumn: number; endColumn: number }): void {
    const book = freshBook();
    if (!book) return;
    if (contDiEdited(book, { row: range.startRow + 1, col: range.startColumn + 1, row2: range.endRow + 1, col2: range.endColumn + 1 })) {
      host.current?.applyWrites(book.writes);
    }
  }

  function macroContext(created: WorkbookHost): MacroContext {
    return {
      host: created,
      book: freshBook,
      notify: setMessage,
      openTool: (section) => setToolRequest({ section, token: Date.now() }),
      openDlEntry: (start, book) => setDlEntry({ start, book }),
      openTargets: setTargets,
      openPipeTally: (note) => refreshPipe(note),
      namedCell,
      jobFolder,
      redrawGauge: () => refreshGauge(created.snapshot()),
      refreshOpenCap,
      redrawAfterSurvey: async () => {
        await created.flushCalculation();
        const book = freshBook();
        if (book) {
          refreshSlideComments(book, false);
          highlightNewestSurvey(book);
          if (book.writes.length > 0) created.applyWrites(book.writes);
        }
        refreshGauge(created.snapshot());
        await created.flushCalculation();
      },
    };
  }

  async function runButton(button: MacroButton) {
    const created = host.current;
    if (!created || busy) return;
    setBusy(true);
    try {
      await runMacro(button, macroContext(created));
    } catch (error) {
      window.alert(`${button.macro} failed: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      setBusy(false);
    }
  }

  function saveDlEntry(start: DlEntryStart, vals: Partial<Record<number, string | number | null>>): string | null {
    const book = freshBook();
    if (!book || !host.current) return "The workbook is not open.";
    try {
      writeEntry(book, start.sheet, start.row, vals);
      host.current.applyWrites(book.writes);
    } catch (error) {
      return error instanceof Error ? error.message : String(error);
    }
    setDlEntry(null);
    setMessage(`DL Tracker row ${start.row} saved.`);
    return null;
  }

  /** frmTargets Save: TargetEditorStore per row, then TargetEditorCommit (SyncPlanTargetWindow). */
  function saveTargets(edits: TargetEdit[]): string | null {
    const book = freshBook();
    if (!book || !host.current) return "The workbook is not open.";
    for (const edit of edits) {
      const message = storeTarget(book, edit.planRow, edit.name, edit.md, edit.inc, edit.azm, edit.tvd);
      if (message) {
        host.current.applyWrites(book.writes);
        return message;
      }
    }
    syncPlanTargetWindow(book);
    host.current.applyWrites(book.writes);
    setTargets(null);
    setMessage("Targets saved. Slidesheet T2:Y5 follows the edited plan targets.");
    return null;
  }

  function refreshPipe(status: string, cleared = false) {
    const book = freshBook();
    if (!book) return;
    const view = readTally(book);
    setPipe({ view: cleared ? { ...view, lengths: [], depths: [] } : view, status });
  }

  async function pipeStep(step: (created: WorkbookHost) => Promise<void>) {
    const created = host.current;
    if (!created || busy) return;
    setBusy(true);
    try {
      await step(created);
    } finally {
      setBusy(false);
    }
  }

  async function writePipeSettings(created: WorkbookHost, stands: string, split: string) {
    const book = freshBook();
    if (!book) return;
    writeSettings(book, stands, split);
    created.applyWrites(book.writes);
    await created.flushCalculation();
  }

  const pipeHandlers = {
    onImport: () => void pipeStep(async (created) => {
      const note = await importTallyCsv(macroContext(created));
      refreshPipe(note || (pipe?.status ?? ""));
    }),
    onClear: () => void pipeStep(async (created) => {
      const book = freshBook();
      if (!book) return;
      clearTallyInputs(book);
      created.applyWrites(book.writes);
      await created.flushCalculation();
      refreshPipe("Cleared pipe tally inputs.", true);
    }),
    onAdd: (stands: string, split: string) => void pipeStep(async (created) => {
      await writePipeSettings(created, stands, split);
      const book = freshBook();
      if (!book || readTally(book).depths.length === 0) {
        window.alert("No tally bit depths to add. Paste pipe lengths first.");
        refreshPipe(pipe?.status ?? "");
        return;
      }
      await appendTally(macroContext(created));
      refreshPipe("Append finished (sheet tally cleared).");
    }),
    onRefresh: () => refreshPipe("Refreshed from sheet."),
    onSettings: (stands: string, split: string) => void pipeStep(async (created) => {
      await writePipeSettings(created, stands, split);
      refreshPipe(pipe?.status ?? "");
    }),
  };

  /** Next period picked in Data!H4:H22: fill the row as MDL_DayRoll does, then ask for depth and hours. */
  function nextDayRoll() {
    while (!dayRollOpen.current && dayRollQueue.current.length > 0) {
      const row = dayRollQueue.current.shift()!;
      const book = host.current?.snapshot() as SnapshotBook | null;
      const started = book ? startDayRoll(book, row) : null;
      if (!started) continue;
      writeData(started.writes);
      dayRollOpen.current = true;
      setDayRoll(started.context);
    }
  }

  function closeDayRoll() {
    dayRollOpen.current = false;
    setDayRoll(null);
    window.setTimeout(nextDayRoll, 0);
  }

  function applyDay(context: DayRollContext, inputs: DayRollInputs): string | null {
    const book = host.current?.snapshot() as SnapshotBook | null;
    if (!book) return "The workbook is not open.";
    const result = applyDayRoll(book, context, inputs);
    if ("error" in result) return result.error;
    writeData(result.writes);
    setMessage(`Data row ${context.row}: End Depth and hours written, slide metres re-placed.`);
    closeDayRoll();
    return null;
  }

  /** The job now open has no workbook: the last job's sheet, buttons, gauge and handoff state leave the screen with it. */
  function closeBook() {
    host.current?.close();
    handoffBlocks.current = [];
    setBookOpen(false);
    setButtons([]);
    setGauge(null);
    setSheetName("");
    setOcJobId("");
    setShiftReady(false);
    setPendingHandoff("");
    setDriftFile("");
    setSyncReview(null);
  }

  async function bind(next: StoredDirectory, label: string, fileName: string): Promise<string> {
    directory.current = next;
    snapshotFileRef.current = fileName;
    // An import left waiting belongs to the folder that was open before.
    setStaged(null);
    const registry = await openRegistry();
    const jobId = appJobIdOf(label);
    const previous = await registry.get(jobId);
    await registry.put({ jobId, label, fileName, lastOpened: new Date().toISOString(), handle: previous?.handle });
    const opened = await openJob({ jobId, label, fileName, directory: next, registry });
    session.current?.dispose();
    session.current = opened.session;
    opened.session.onStateChange = (nextState, error) => {
      if (session.current !== opened.session) return;
      setStatus(nextState);
      if (error !== undefined) {
        setMessage(folderGone(error) ? missingFolderText(label) : error instanceof Error ? error.message : String(error));
      }
    };
    setRecovery(opened.result.recovery);
    setStatus(opened.result.state);
    setBackups(opened.result.recovery.kind === "corrupt" || opened.result.recovery.kind === "future" ? opened.result.recovery.backups.length : 0);
    if (opened.result.envelope) {
      const withButtons = await refreshFromSource(next, opened.result.envelope, fileName);
      host.current?.load(withButtons.workbook);
      setBookOpen(true);
      refreshGauge(host.current?.snapshot());
      setOcJobId(workbookOcJobId());
      setButtons(buttonsOf(withButtons.sidecar));
      setSheetName(host.current?.activeSheet() ?? "");
      setAuthority(authorityOf(withButtons.sidecar as { authority?: unknown }));
      handoffBlocks.current = Array.isArray(withButtons.handoffBlocks) ? withButtons.handoffBlocks : [];
      setShiftReady(isCanonBook(withButtons.importBaseline));
      const repaired = await repairDamagedSurveyFormulas();
      setMessage(`${label} opened from the folder snapshot.${repaired}`);
      // Check Excel drift BEFORE autosyncOpenJob mutates Data/Setup/Slidesheet.
      const hasConflictPrompt = await refreshHandoff(next, false);
      if (!hasConflictPrompt) void autosyncOpenJob(jobId);
    } else {
      closeBook();
      setMessage(opened.result.recovery.kind === "none"
        ? `${label} has no Slide Sheet open yet. Put the job's .xlsm in its folder.`
        : recoveryMessage(opened.result.recovery));
    }
    setJobs(await registry.list());
    setJobId(jobId);
    setPickedJob(null);
    return jobId;
  }

  /** Backfill buttons and cell formatting from the folder .xlsm (read-only) when the snapshot predates them. */
  async function refreshFromSource(folder: StoredDirectory, envelope: PhxEnvelope, fileName = ""): Promise<PhxEnvelope> {
    const sidecar = envelope.sidecar && typeof envelope.sidecar === "object" ? envelope.sidecar as Record<string, unknown> : {};
    const needButtons = buttonsOf(sidecar).length === 0;
    const needStyles = sidecar.styleVersion !== STYLE_VERSION;
    if (!needButtons && !needStyles) return envelope;
    const xlsm = matchingWorkbook(await folder.list(), fileName || snapshotFileRef.current, ocJobId);
    if (!xlsm) return envelope;
    const file = await folder.getFile(xlsm, false);
    const bytes = file ? await file.read() : null;
    if (!bytes) return envelope;
    let workbook = envelope.workbook;
    const patch: Record<string, unknown> = {};
    if (needButtons) {
      const macroButtons = macroButtonsFromWorkbook(bytes);
      if (macroButtons.length > 0) patch.macroButtons = macroButtons;
    }
    if (needStyles && workbook && typeof workbook === "object") {
      setMessage(`Restoring cell formatting from ${xlsm}. Values stay as they are.`);
      const fresh = await readWorkbook(bytes);
      workbook = restyleWorkbook(workbook, fresh.workbook);
      patch.styleVersion = STYLE_VERSION;
    }
    if (Object.keys(patch).length === 0) return envelope;
    const next = stamp(workbook, { ...sidecar, ...patch }, { hash: envelope.sourceWorkbookHash, baseline: envelope.importBaseline, blocks: envelope.handoffBlocks ?? [] });
    session.current?.noteEdit(next);
    return next;
  }

  async function openKnown(nextJobId: string) {
    setPickedJob(null);
    setJobId(nextJobId);
    const registry = await openRegistry();
    const record = (await registry.list()).find((job) => job.jobId === nextJobId);
    if (!record) return;
    const folder = (await reopenStoredFolder(registry, nextJobId)) ?? adapterFolder.current;
    if (!folder) {
      setRecovery({ kind: "permission", jobId: nextJobId, label: record.label });
      setMessage(`${record.label} needs the folder grant again. The mirror is still in this browser.`);
      return;
    }
    let permission = await folder.queryPermission();
    directory.current = folder;
    if (permission === "prompt") {
      // Choosing the job in the list is the click that lets the browser ask for the folder again.
      // Chrome often keeps the grant between sessions; Edge asks after a restart.
      permission = await folder.requestPermission().catch(() => "prompt" as const);
    }
    if (permission !== "granted") {
      setRecovery({ kind: "permission", jobId: nextJobId, label: record.label });
      setStatus("permission-blocked");
      setMessage(`${record.label} needs the folder grant again. The mirror is still in this browser.`);
      return;
    }
    try {
      // The folder is read again on every open, so a workbook put there since is picked up.
      const names = await folder.list();
      const found = folderWorkbooks(names, record.fileName);
      if (found.snapshot && found.snapshot !== record.fileName) {
        // The workbook was imported under its own name after this entry was made for the empty folder.
        const label = snapshotLabel(found.snapshot);
        const keptId = appJobIdOf(label);
        if (keptId !== nextJobId) {
          const [stale, kept] = await Promise.all([registry.get(nextJobId), registry.get(keptId)]);
          if (!kept?.handle) await registry.put({ jobId: keptId, label, fileName: found.snapshot, lastOpened: new Date().toISOString(), handle: stale?.handle });
          await retireEntry(registry, nextJobId, keptId);
        }
        await bind(folder, label, found.snapshot);
      } else {
        // Before the job opens: the FieldCap job linked to a stale entry has to follow this one, or autosync pulls the job the workbook was copied from.
        await retireStaleEntries(registry, nextJobId, (await registry.get(nextJobId))?.handle, names);
        await bind(folder, record.label, record.fileName);
      }
      setFolderXlsm(found.snapshot ? "" : found.xlsm ?? "");
      if (!found.snapshot && found.xlsm) await stageFile(folder, found.xlsm);
    } catch (error) {
      setStatus("save-error");
      setMessage(folderGone(error) ? missingFolderText(record.label) : error instanceof Error ? error.message : String(error));
    }
  }

  /** One folder under two job-list entries: `fromId` goes, and the FieldCap jobs linked to it follow `toId`. */
  async function retireEntry(registry: JobRegistry, fromId: string, toId: string) {
    await moveJobLinks(fromId, toId);
    await registry.remove(fromId);
    setLinks(await jobMap());
  }

  /** The folder being opened may be listed a second time from before its workbook arrived. That entry goes. */
  async function retireStaleEntries(registry: JobRegistry, keptId: string, handle: unknown, names: string[]) {
    const entries = await Promise.all((await registry.list()).map(async (record) => ({ ...record, handle: (await registry.get(record.jobId))?.handle })));
    for (const staleId of await staleEntriesFor(entries, keptId, handle, names)) await retireEntry(registry, staleId, keptId);
  }

  /** Ask for a job folder and open it. Returns its app job id, or null when the picker was cancelled. */
  async function addFolder(): Promise<string | null> {
    try {
      const registry = await openRegistry();
      const adapter = adapterFolder.current;
      const next = adapter ?? await grantJobFolder(pickedDirectoryHandle, registry, {
        jobId: "pending-folder",
        label: "Job",
        fileName: "pending.phx.json",
      });
      // A new job folder with no workbook in it yet is named after the folder itself.
      let folderName = "Job";
      const pending = adapter ? null : await registry.get("pending-folder");
      if (!adapter) {
        await registry.remove("pending-folder");
        directory.current = next;
        folderName = String((pending?.handle as { name?: unknown } | undefined)?.name ?? "").trim() || folderName;
      }
      const names = await next.list();
      const { snapshot, xlsm } = folderWorkbooks(names);
      const label = (snapshot && snapshotLabel(snapshot)) || xlsm?.replace(/\.xlsm$/i, "") || folderName;
      const fileName = snapshot ?? `${label}.phx.json`;
      if (!adapter) {
        const keptId = appJobIdOf(label);
        await registry.put({ jobId: keptId, label, fileName, lastOpened: new Date().toISOString(), handle: pending?.handle });
        await retireStaleEntries(registry, keptId, pending?.handle, names);
      }
      const opened = await bind(next, label, fileName);
      setFolderXlsm(snapshot ? "" : xlsm ?? "");
      if (!snapshot && xlsm) await stageFile(next, xlsm);
      return opened;
    } catch (error) {
      if (error instanceof DOMException && error.name === "AbortError") return null;
      setStatus("save-error");
      setMessage(error instanceof Error ? error.message : String(error));
      return null;
    }
  }

  async function stageFile(folder: StoredDirectory, fileName: string) {
    setMessage(`Reading ${fileName}`);
    const file = await folder.getFile(fileName, false);
    const bytes = file ? await file.read() : null;
    if (!bytes) throw new Error(`${fileName} could not be read.`);
    const mapped = await readWorkbook(bytes);
    setStaged({ mapped, fileName });
    setMessage(`Preview ready for ${fileName}. Continue keeps this evaluation copy. Cancel leaves the current preview.`);
  }

  async function continueImport() {
    if (!staged || !session.current || !host.current) return;
    host.current.setUdfNames(staged.mapped.importLog.unknownFunctions);
    host.current.load(staged.mapped.workbook);
    const snap = host.current.snapshot();
    if (!snap) return;
    setBookOpen(true);
    setFolderXlsm("");
    handoffBlocks.current = [];
    const envelope: PhxEnvelope = stamp(snap, staged.mapped.sidecar, {
      hash: staged.mapped.sourceSha256,
      baseline: canonicalize(snap),
      blocks: [],
    });
    try {
      await session.current.replace(envelope, { backup: true });
    } catch (error) {
      const label = session.current.label || jobs.find((job) => job.jobId === jobId)?.label || "this job";
      setStatus(session.current.state);
      setMessage(folderGone(error) ? missingFolderText(label) : error instanceof Error ? error.message : String(error));
      return;
    }
    refreshGauge(envelope.workbook);
    setOcJobId(workbookOcJobId());
    setStatus(session.current.state);
    setButtons(buttonsOf(envelope.sidecar));
    setSheetName(host.current.activeSheet());
    setStaged(null);
    setShiftReady(true);
    setDriftFile("");
    setSyncReview(null);
    dismissedDrift.current = "";
    machinePreferBrowser.current = false;
    setPendingHandoff("");
    setHandoffNote("Browser shift started. Excel stays the authority until Export handoff is applied there.");
    const repaired = await repairDamagedSurveyFormulas();
    setMessage(`Import saved as a disposable preview. Excel is still the authority.${safetyCopyNote()}${repaired}`);
  }

  /** The folder save went through but the browser safety copy did not: said once, beside the save it belongs to. */
  function safetyCopyNote(): string {
    const reason = session.current?.mirrorError ?? "";
    return reason ? ` The browser safety copy was not written (${reason}), so edits are only safe once the folder has them.` : "";
  }

  async function acceptNative() {
    const current = session.current?.current;
    const snap = host.current?.snapshot();
    if (!current || !snap || !session.current) return;
    const sidecar = current.sidecar && typeof current.sidecar === "object" ? current.sidecar as Record<string, unknown> : {};
    const envelope: PhxEnvelope = stamp(snap, { ...sidecar, authority: "native", excelFallback: "locked-on-disk" });
    session.current.noteEdit(envelope);
    await session.current.flushFolder();
    setAuthority("native");
    setStatus(session.current.state);
    setMessage("This folder snapshot is now the job record. The Excel file was not changed. Mark that file read-only on disk.");
  }

  function syncJobLabel(browserBook: CanonBook, excelBook: CanonBook): string {
    const cellVal = (book: CanonBook, sheetName: string, a1: string) =>
      book.sheets.find((sheet) => sheet.name === sheetName)?.cells[a1]?.v ?? "";
    const jobNo = cellVal(browserBook, "Setup", "D4") || cellVal(excelBook, "Setup", "D4") || workbookOcJobId();
    const excelBha = cellVal(excelBook, "Data", "H3");
    const browserBha = cellVal(browserBook, "Data", "H3");
    const parts: string[] = [];
    if (jobNo) parts.push(`Job ${jobNo}`);
    if (excelBha || browserBha) {
      parts.push(excelBha === browserBha
        ? `BHA #${excelBha}`
        : `Excel BHA #${excelBha || "?"} vs Browser BHA #${browserBha || "?"}`);
    }
    return parts.join(" · ");
  }

  async function commitInbound(review: SyncReview, picks: Record<string, "excel" | "browser">) {
    const created = host.current;
    const current = session.current?.current;
    if (!created || !current || !session.current) return;
    const writes = review.merge.fromExcel.flatMap((op) => opWrites(op));
    for (const conflict of review.merge.conflicts) {
      const key = `${conflict.sheet}!${conflict.a1}`;
      if ((picks[key] ?? "excel") === "excel") writes.push(...scalarWrites(conflict.sheet, conflict.a1, conflict.excel));
    }
    if (writes.length > 0) created.applyWrites(writes);
    await created.flushCalculation();
    if (writes.some((write) => write.sheet === "Data")) {
      const book = freshBook();
      if (book) {
        toolHoursSync(book);
        if (book.writes.length > 0) {
          created.applyWrites(book.writes);
          await created.flushCalculation();
        }
      }
    }
    const snap = created.snapshot();
    if (!snap) return;
    refreshGauge(snap);
    const sidecar = current.sidecar && typeof current.sidecar === "object" ? current.sidecar : {};
    session.current.noteEdit(stamp(snap, sidecar, { hash: review.hash, baseline: review.excel }));
    await session.current.flushFolder();
    const kept = review.merge.keepBrowser.length + review.merge.conflicts.filter((conflict) => (picks[`${conflict.sheet}!${conflict.a1}`] ?? "excel") === "browser").length;
    machinePreferBrowser.current = false;
    dismissedDrift.current = "";
    setDriftFile("");
    setSyncReview(null);
    setShiftReady(true);
    setStatus(session.current.state);
    setHandoffNote(kept > 0
      ? `Synced ${review.fileName}. ${writes.length} cell(s) came from Excel. Setup & Slidesheet stay from browser (${kept} ready for Export handoff).`
      : `Synced ${review.fileName}. ${writes.length} cell(s) came from Excel.`);
  }

  /** Returns true when a conflict review dialog is open for Data / Cont DI / DL Tracker. */
  async function refreshHandoff(folder: StoredDirectory, force = false): Promise<boolean> {
    const handle = await jobFolder();
    if (handle) {
      const pending = await pendingHandoffs(handle).catch(() => []);
      setPendingHandoff(pending[0] ?? "");
      const receipt = await latestReceipt(handle).catch(() => null);
      if (receipt?.status && !force) {
        setHandoffNote(receipt.status === "applied"
          ? `Excel applied handoff ${receipt.handoffId}.`
          : `Excel reported ${receipt.status} for handoff ${receipt.handoffId}. ${receipt.message}`);
      }
    }
    const current = session.current?.current;
    if (!current?.sourceWorkbookHash || !isCanonBook(current.importBaseline) || !host.current) {
      setDriftFile("");
      return false;
    }
    const xlsm = matchingWorkbook(await folder.list(), snapshotFileRef.current, workbookOcJobId() || ocJobId);
    if (!xlsm) {
      setDriftFile("");
      return false;
    }
    const file = await folder.getFile(xlsm, false);
    const bytes = file ? await file.read() : null;
    if (!bytes) return false;
    const hash = await sha256(bytes);
    if (hash === current.sourceWorkbookHash) {
      setDriftFile("");
      dismissedDrift.current = "";
      if (force) setHandoffNote("The saved workbook matches this shift. Save Excel first if you just edited it.");
      return false;
    }
    setDriftFile(xlsm);
    if (!force && (dismissedDrift.current === hash || syncReview)) return Boolean(syncReview);
    if ((busy || session.current?.folderUnsafe()) && !force) return false;
    if (inboundLock.current) return Boolean(syncReview);
    const snap = host.current.snapshot();
    if (!snap) return false;
    inboundLock.current = true;
    try {
      const mapped = await readWorkbook(bytes);
      const browserCanon = canonicalize(snap);
      const excel = canonicalize(mapped.workbook);
      const merge = mergeVisibleUserData(current.importBaseline, browserCanon, excel);
      const review: SyncReview = {
        fileName: xlsm,
        jobLabel: syncJobLabel(browserCanon, excel),
        hash,
        excel,
        merge,
        preferBrowserMachine: machinePreferBrowser.current,
      };
      // Setup & Slidesheet are strictly Browser-authoritative. Auto-merge whenever Data/Cont DI/DL Tracker have zero conflicts.
      if (merge.conflicts.length === 0) {
        await commitInbound(review, {});
        return false;
      }
      setSyncReview(review);
      return true;
    } catch (error) {
      setHandoffNote(error instanceof Error ? error.message : String(error));
      return false;
    } finally {
      inboundLock.current = false;
    }
  }
  refreshHandoffRef.current = refreshHandoff;

  async function exportHandoff() {
    const current = session.current?.current;
    const snap = host.current?.snapshot();
    const folder = await jobFolder();
    if (!current || !snap || !folder || !directory.current) {
      setHandoffNote("Open a job folder and import its workbook before exporting a handoff.");
      return;
    }
    if (!isCanonBook(current.importBaseline)) {
      setHandoffNote("Import the workbook to start a browser shift. Excel stays the authority.");
      return;
    }
    if (driftFile) {
      setHandoffNote("Excel has changed. Save the workbook in Excel, then Sync with Excel, before exporting.");
      return;
    }
    if (session.current?.folderUnsafe()) {
      setHandoffNote("The snapshot is still saving.");
      return;
    }
    const pending = await pendingHandoffs(folder);
    if (pending.length > 0) {
      setPendingHandoff(pending[0] ?? "");
      setHandoffNote(`Handoff ${pending[0]} is already waiting for Excel.`);
      return;
    }
    const ops = diffVisibleUserData(current.importBaseline, canonicalize(snap));
    if (ops.length === 0) {
      setHandoffNote("Nothing in this shift needs to go back to Excel.");
      return;
    }
    const xlsm = matchingWorkbook(await directory.current.list(), snapshotFileRef.current, workbookOcJobId() || ocJobId) ?? "";
    const token = workbookToken(xlsm);
    if (!token) {
      setHandoffNote("The workbook file name has no job token Excel can check.");
      return;
    }
    setBusy(true);
    setHandoffNote("Preparing the handoff.");
    try {
      const id = `handoff-${new Date().toISOString().replace(/[:.]/g, "-")}`;
      const files = await buildBundleFiles({
        jobNumber: token,
        sourceWorkbookHash: current.sourceWorkbookHash,
        createdAt: new Date().toISOString(),
        exporter: "browser shift",
        mode: "visible-data",
        ops,
      });
      await commitHandoff(folder, id, files);
      setPendingHandoff(id);
      setHandoffNote(`Handoff ${id} is ready (${ops.length} cells). Open the workbook in Excel and choose Sync from Browser.`);
    } catch (error) {
      setPendingHandoff("");
      setHandoffNote(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  async function saveNow() {
    try {
      await session.current?.flushFolder();
      setStatus(session.current?.state ?? "folder-safe");
      setMessage(`Folder snapshot saved.${safetyCopyNote()}`);
    } catch (error) {
      const label = session.current?.label || jobs.find((job) => job.jobId === jobId)?.label || "this job";
      setStatus(session.current?.state ?? "save-error");
      setMessage(folderGone(error) ? missingFolderText(label) : error instanceof Error ? error.message : String(error));
    }
  }

  async function reconnect() {
    if (!directory.current || recovery.kind !== "permission") return;
    const granted = await directory.current.requestPermission();
    if (granted !== "granted") {
      setMessage("Folder permission is still blocked. The mirror was kept.");
      return;
    }
    await openKnown(recovery.jobId);
  }

  async function keepFolder() {
    if (recovery.kind !== "mirror-newer" || !session.current) return;
    const registry = await openRegistry();
    await registry.removeMirror(recovery.jobId);
    setRecovery({ kind: "none" });
    setMessage("Kept the folder snapshot and marked the mirror resolved.");
  }

  async function restoreMirror() {
    if (recovery.kind !== "mirror-newer") return;
    const registry = await openRegistry();
    const mirror = await registry.getMirror(recovery.jobId);
    if (!mirror || !session.current || !host.current) return;
    const envelope = parseEnvelope(await mirrorSnapshot(mirror));
    host.current.load(envelope.workbook);
    setBookOpen(true);
    try {
      await session.current.replace(envelope, { backup: true });
    } catch (error) {
      const label = session.current.label || jobs.find((job) => job.jobId === jobId)?.label || "this job";
      setStatus(session.current.state);
      setMessage(folderGone(error) ? missingFolderText(label) : error instanceof Error ? error.message : String(error));
      return;
    }
    setRecovery({ kind: "none" });
    setStatus(session.current.state);
    const repaired = await repairDamagedSurveyFormulas();
    setMessage(`Restored the mirror into the folder.${repaired}`);
  }

  return (
    <div className="shell">
      <div className="toolbar">
        <strong>PHX preview</strong>
        <span className="stamp">{PREVIEW_VERSION}</span>
        <div className="view-switch" role="group" aria-label="View">
          <button type="button" aria-pressed={surface === "sheets"} onClick={() => setSurface("sheets")}>Sheets</button>
          <button type="button" aria-pressed={surface === "wells"} onClick={() => setSurface("wells")}>3D Wells</button>
        </div>
        <JobPicker
          jobs={fieldCapList}
          jobsError={fieldCapError}
          folders={jobs}
          links={links}
          appJobId={jobId}
          picked={pickedJob}
          busy={busy}
          onOpen={(nextJobId) => void openKnown(nextJobId)}
          onPick={(job) => {
            setPickedJob(job);
            const open = jobs.find((folder) => folder.jobId === jobId)?.label;
            setMessage(`Job ${job.jobNumber} has no folder in this browser yet. Choose its job folder; Cancel keeps the job selected, and Pull from FieldCap and Publish to StrapSheets still work without one.${open ? ` The workbook below is still ${open}.` : ""}`);
            // A browser can only open a disk folder after the user picks it once, so the pick starts here, from the same click.
            void chooseFolder(job);
          }}
          onChooseFolder={(job) => void chooseFolder(job)}
        />
        <button type="button" onClick={() => void saveNow()}>Save now</button>
        <button
          type="button"
          aria-pressed={menusShown}
          onClick={() => {
            host.current?.setMenusVisible(!menusShown);
            setMenusShown(!menusShown);
          }}
          title="The spreadsheet menus (Start, Insert, Formulas, Data, View) and the formula bar."
        >
          {menusShown ? "Hide Menus" : "Show Menus"}
        </button>
        <button
          type="button"
          disabled={busy || !shiftReady || status === "saving-folder" || status === "saving-mirror" || pendingHandoff !== "" || driftFile !== ""}
          onClick={() => void exportHandoff()}
          title="Writes a handoff Excel applies the next time this workbook opens. Save Excel and sync first if the file on disk has changed. This does not edit the .xlsm."
        >
          Export handoff
        </button>
        <button
          type="button"
          disabled={busy || !shiftReady}
          onClick={() => {
            const folder = directory.current;
            if (folder) void refreshHandoff(folder, true);
            else setHandoffNote("Open a job folder first. Save the workbook in Excel before syncing.");
          }}
          title="Reads the last saved .xlsm. Save Excel first."
        >
          Sync with Excel
        </button>
        <button
          type="button"
          disabled={busy || !selectedJob}
          onClick={() => void pullFieldCap()}
          title={selectedJob ? `Reads FieldCap job ${selectedJob.jobNumber} only. Writes its OpenCap CSVs into the job folder's OpenCap and refreshes Setup, Costs and the mud motors.` : "Choose a FieldCap job in the job list first."}
        >
          Pull from FieldCap
        </button>
        <button
          type="button"
          disabled={busy || (!selectedJob && !(jobId && ocJobId))}
          onClick={() => void publishStrapSheets()}
          title="Sends this job's pulled inventory and crew to StrapSheets. Signs in with Microsoft or Google the first time."
        >
          Publish to StrapSheets
        </button>
        <Tools
          host={() => host.current}
          book={() => (host.current?.snapshot() as SnapshotBook | null) ?? null}
          authority={authority}
          onAuthority={acceptNative}
          onStatus={setMessage}
          onEmailReport={async () => {
            const created = host.current;
            if (created) await emailReport(macroContext(created));
          }}
          toolRequest={toolRequest}
        />
      </div>
      {recovery.kind === "permission" ? (
        <section className="panel">
          <h1>{recovery.label} needs folder permission</h1>
          <p>The mirror is still stored in this browser. Reconnect grants the folder again.</p>
          <button type="button" onClick={() => void reconnect()}>Reconnect</button>
        </section>
      ) : null}
      {recovery.kind === "corrupt" || recovery.kind === "future" ? (
        <section className="panel">
          <h1>Snapshot needs recovery</h1>
          <p>{recoveryMessage(recovery)} Validated backups: {backups}.</p>
        </section>
      ) : null}
      {recovery.kind === "mirror-newer" ? (
        <section className="panel">
          <h1>A newer browser mirror exists</h1>
          <button type="button" onClick={() => void restoreMirror()}>Restore mirror</button>
          <button type="button" onClick={() => void keepFolder()}>Keep folder</button>
        </section>
      ) : null}
      {driftFile && !syncReview ? (
        <section className="panel">
          <h1>{driftFile} changed in Excel</h1>
          <p>Save the workbook in Excel first. Sync with Excel merges this shift. Export handoff stays off until that merge has run.</p>
          <button type="button" onClick={() => directory.current && void refreshHandoff(directory.current, true)}>Sync with Excel</button>
        </section>
      ) : null}
      {syncReview ? (
        <SyncConflictDialog
          fileName={syncReview.fileName}
          jobLabel={syncReview.jobLabel}
          merge={syncReview.merge}
          preferBrowserMachine={syncReview.preferBrowserMachine}
          onApply={(picks) => void commitInbound(syncReview, picks)}
          onCancel={() => {
            dismissedDrift.current = syncReview.hash;
            setSyncReview(null);
            setHandoffNote("Excel has changed. Export handoff stays off until Sync with Excel merges this shift.");
          }}
        />
      ) : null}
      {staged ? (
        <section className="panel">
          <h1>Start browser shift from {staged.fileName}</h1>
          <p>Excel stays the authority. Continue captures this file as the shift baseline.</p>
          <p className="log">{summarize(staged.mapped.importLog)}</p>
          <button type="button" onClick={() => void continueImport()}>Continue</button>
          <button type="button" onClick={() => setStaged(null)}>Cancel</button>
        </section>
      ) : null}
      {dayRoll ? (
        <DayRollDialog
          key={dayRoll.row}
          context={dayRoll}
          onApply={(inputs) => applyDay(dayRoll, inputs)}
          onCancel={closeDayRoll}
        />
      ) : null}
      {dlEntry ? (
        <DlEntryDialog
          key={dlEntry.start.row}
          book={dlEntry.book}
          start={dlEntry.start}
          onSave={(vals) => saveDlEntry(dlEntry.start, vals)}
          onCancel={() => setDlEntry(null)}
        />
      ) : null}
      {targets ? <TargetsDialog targets={targets} onSave={saveTargets} onCancel={() => setTargets(null)} /> : null}
      {pipe ? (
        <PipeTallyDialog
          view={pipe.view}
          status={pipe.status}
          busy={busy}
          {...pipeHandlers}
          onClose={() => setPipe(null)}
        />
      ) : null}
      <div className="workspace">
        <div className="grid-wrap">
          <div ref={grid} className="grid" />
          <SheetOverlay host={workbookHost} sheet={sheetName} buttons={buttons} gauge={gauge} busy={busy} onMacro={(button) => void runButton(button)} />
          <WellsView host={workbookHost} open={bookOpen} active={surface === "wells"} />
          {!bookOpen && recovery.kind === "none" ? (
            <NoWorkbook
              reason={noWorkbookReason({
                appJobId: jobId,
                label: jobs.find((job) => job.jobId === jobId)?.label ?? "",
                pickedJobNumber: pickedJob?.jobNumber ?? "",
                stagedFile: staged?.fileName ?? "",
                folderXlsm,
              })}
              busy={busy}
              onOpenWorkbook={(xlsm) => directory.current && void stageFile(directory.current, xlsm).catch((error: unknown) => setMessage(error instanceof Error ? error.message : String(error)))}
              onCheckFolder={() => void openKnown(jobId)}
            />
          ) : null}
        </div>
        <div id={TOOLS_DOCK_ID} className="dock" />
      </div>
      <div className="status" role="status" aria-live="polite">
        <span>{status}</span>
        <span>{message}</span>
        <span>{handoffNote}</span>
        <span>{backups === 1 ? "1 restore point" : `${backups} restore points`}</span>
      </div>
    </div>
  );
}

function summarize(entry: ImportLog): string {
  return [
    `Sheets ${entry.sheetCount}`,
    `Formulas ${entry.formulaCount}`,
    `Shared ${entry.sharedFormulaCount}`,
    `Arrays ${entry.arrayFormulaCount}`,
    `Names ${entry.supportedNameCount}/${entry.definedNameCount}`,
    `UDF names ${entry.unknownFunctions.join(", ") || "none"}`,
    `External refs rejected ${entry.rejectedExternalRefs.length}`,
    `Macro buttons ${entry.macroButtonCount}`,
  ].join("\n");
}

function buttonsOf(sidecar: unknown): MacroButton[] {
  if (!sidecar || typeof sidecar !== "object" || !("macroButtons" in sidecar)) return [];
  const buttons = (sidecar as { macroButtons?: unknown }).macroButtons;
  return Array.isArray(buttons) ? buttons as MacroButton[] : [];
}

function recoveryMessage(recovery: Recovery): string {
  if (recovery.kind === "permission") return `${recovery.label} is waiting for folder permission.`;
  if (recovery.kind === "future") return "This snapshot was written by a newer preview and was left untouched.";
  if (recovery.kind === "corrupt") return "The folder snapshot could not be read and was left untouched.";
  if (recovery.kind === "mirror-newer") return "The browser mirror is newer than the folder snapshot.";
  return "";
}
