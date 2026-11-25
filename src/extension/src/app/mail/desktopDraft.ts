import { toDraftMime, type MailDraft, type TransportResult } from "@phx/mail";
import { archiveFileName } from "./archive";

export interface DownloadChangeDelta {
  id: number;
  state?: { current?: string };
  error?: { current?: string };
}

export interface DownloadsApi {
  download(options: {
    url: string;
    filename: string;
    saveAs: boolean;
    conflictAction?: "uniquify" | "overwrite" | "prompt";
  }): Promise<number>;
  search(query: { id: number }): Promise<Array<{ id: number; state?: string; error?: string }>>;
  open(downloadId: number): Promise<void>;
  onChanged: {
    addListener(listener: (delta: DownloadChangeDelta) => void): void;
    removeListener(listener: (delta: DownloadChangeDelta) => void): void;
  };
}

export interface DesktopDraftDeps {
  downloads?: DownloadsApi;
  createObjectUrl?: (blob: Blob) => string;
  revokeObjectUrl?: (url: string) => void;
  now?: () => Date;
  onOpening?: () => void;
  timeoutMs?: number;
}

export interface DesktopDraftOutcome {
  result: TransportResult;
  downloadId: number | null;
}

function chromeDownloads(): DownloadsApi {
  return {
    download(options) {
      return new Promise((resolve, reject) => {
        chrome.downloads.download(options, (id) => {
          const err = chrome.runtime.lastError?.message;
          if (err || typeof id !== "number") reject(new Error(err || "Could not save the Outlook draft file."));
          else resolve(id);
        });
      });
    },
    search(query) {
      return new Promise((resolve, reject) => {
        chrome.downloads.search(query, (items) => {
          const err = chrome.runtime.lastError?.message;
          if (err) reject(new Error(err));
          else resolve(items ?? []);
        });
      });
    },
    open(downloadId) {
      return new Promise((resolve, reject) => {
        try {
          const maybePromise = chrome.downloads.open(downloadId) as unknown as Promise<void> | undefined;
          if (maybePromise && typeof maybePromise.then === "function") {
            maybePromise.then(() => resolve(), (error) => reject(error instanceof Error ? error : new Error(String(error))));
            return;
          }
          const err = chrome.runtime.lastError?.message;
          if (err) reject(new Error(err));
          else resolve();
        } catch (error) {
          reject(error instanceof Error ? error : new Error(String(error)));
        }
      });
    },
    onChanged: chrome.downloads.onChanged,
  };
}

function waitForComplete(downloads: DownloadsApi, downloadId: number, timeoutMs: number): Promise<void> {
  return new Promise((resolve, reject) => {
    let settled = false;
    const finish = (error?: Error) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      downloads.onChanged.removeListener(onChange);
      if (error) reject(error);
      else resolve();
    };
    const onChange = (delta: DownloadChangeDelta) => {
      if (delta.id !== downloadId) return;
      if (delta.state?.current === "complete") finish();
      else if (delta.state?.current === "interrupted") {
        finish(new Error(delta.error?.current ? `Draft download was interrupted (${delta.error.current}).` : "Draft download was interrupted."));
      }
    };
    const timer = setTimeout(() => finish(new Error("Timed out saving the Outlook draft file.")), timeoutMs);
    downloads.onChanged.addListener(onChange);
    void downloads.search({ id: downloadId }).then((items) => {
      const item = items[0];
      if (!item) return;
      if (item.state === "complete") finish();
      else if (item.state === "interrupted") {
        finish(new Error(item.error ? `Draft download was interrupted (${item.error}).` : "Draft download was interrupted."));
      }
    }).catch(() => undefined);
  });
}

/**
 * Saves an editable X-Unsent: 1 .eml draft (To/Cc, Subject, HTML table, inline corridor PNG, and all attachments)
 * and asks Windows to open it directly in the default Outlook client (Classic Outlook or New Outlook).
 */
export async function openDesktopDraft(draft: MailDraft, deps: DesktopDraftDeps = {}): Promise<DesktopDraftOutcome> {
  const downloads = deps.downloads ?? chromeDownloads();
  const createUrl = deps.createObjectUrl ?? ((blob: Blob) => URL.createObjectURL(blob));
  const revokeUrl = deps.revokeObjectUrl ?? ((url: string) => URL.revokeObjectURL(url));
  const now = deps.now ?? (() => new Date());
  const timeoutMs = deps.timeoutMs ?? 15_000;

  const eml = toDraftMime(draft);
  const blob = new Blob([eml], { type: "message/rfc822" });
  const url = createUrl(blob);
  const filename = `PHX-Reports/${archiveFileName(now())}`;
  let downloadId: number;
  try {
    downloadId = await downloads.download({ url, filename, saveAs: false, conflictAction: "uniquify" });
    await waitForComplete(downloads, downloadId, timeoutMs);
  } catch (error) {
    return {
      downloadId: null,
      result: { status: "no-side-effect-failure", error: error instanceof Error ? error.message : "Could not save the Outlook draft file." },
    };
  } finally {
    revokeUrl(url);
  }

  deps.onOpening?.();
  try {
    await downloads.open(downloadId);
    return {
      downloadId,
      result: { status: "created-and-opened", id: String(downloadId), location: "Outlook" },
    };
  } catch (error) {
    return {
      downloadId,
      result: {
        status: "created-not-opened",
        id: String(downloadId),
        location: "Outlook",
        error: error instanceof Error ? error.message : "Chrome requires a click to open the draft.",
      },
    };
  }
}

export async function openDownloadedDraft(downloadId: number, downloads: DownloadsApi = chromeDownloads()): Promise<TransportResult> {
  try {
    await downloads.open(downloadId);
    return { status: "created-and-opened", id: String(downloadId), location: "Outlook" };
  } catch (error) {
    return {
      status: "created-not-opened",
      id: String(downloadId),
      location: "Outlook",
      error: error instanceof Error ? error.message : "Could not open the draft in Outlook.",
    };
  }
}
