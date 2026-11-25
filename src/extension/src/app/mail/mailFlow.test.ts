import { afterEach, describe, expect, it, vi } from "vitest";
import { createDraft, type MailDraft, type TransportResult } from "@phx/mail";
import { composeOnPort, probeClassic, type HostPort, type HostReply } from "./classicHost";
import { openDesktopDraft, openDownloadedDraft, type DownloadChangeDelta, type DownloadsApi } from "./desktopDraft";
import { chooseSuggestion, looksLikeClassic } from "./detect";
import { mailtoUrl } from "./genericMail";
import { createGmailDraft } from "./gmailDraft";
import { createMicrosoftDraft } from "./microsoftDraft";
import { deliverReport, rememberRetry, retryOpen, type ReportTransport } from "./pipeline";
import { clearSessionToken, readSessionToken, saveSessionToken, type KeyValue, type SessionToken } from "./platform";

function draft(): MailDraft {
  return createDraft({
    to: ["a@example.com"],
    cc: [],
    subject: "Daily Drilling Summary - Job",
    html: "<p>Table</p>",
    text: "Table",
    inline: [],
    attachments: [{ name: "note.txt", contentType: "text/plain", bytes: Uint8Array.of(97) }],
    missing: [],
  });
}

function memory(): KeyValue & { values: Record<string, unknown> } {
  const values: Record<string, unknown> = {};
  return {
    values,
    async get(keys) {
      const found: Record<string, unknown> = {};
      for (const key of keys) found[key] = values[key];
      return found;
    },
    async set(next) { Object.assign(values, next); },
    async remove(keys) { for (const key of keys) delete values[key]; },
  };
}

function scripted(replies: HostReply[]): { port: HostPort; sent: unknown[] } {
  const sent: unknown[] = [];
  let onMessage: ((message: HostReply) => void) | null = null;
  const disconnects = new Set<() => void>();
  const port: HostPort = {
    postMessage(message) {
      sent.push(message);
      const reply = replies.shift();
      if (reply) onMessage?.(reply);
    },
    onMessage: {
      addListener(listener) { onMessage = listener; },
      removeListener() { onMessage = null; },
    },
    onDisconnect: {
      addListener(listener) { disconnects.add(listener); },
      removeListener(listener) { disconnects.delete(listener); },
    },
    disconnect() {},
  };
  return { port, sent };
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("Classic Outlook helper", () => {
  it("consumes and reports Chrome's native-host disconnect error", async () => {
    let lastErrorReads = 0;
    vi.stubGlobal("chrome", {
      runtime: {
        get lastError() {
          lastErrorReads += 1;
          return { message: "Specified native messaging host not found." };
        },
      },
    });
    let disconnect: (() => void) | null = null;
    const port: HostPort = {
      postMessage() { disconnect?.(); },
      onMessage: {
        addListener() {},
        removeListener() {},
      },
      onDisconnect: {
        addListener(listener) { disconnect = listener; },
        removeListener() { disconnect = null; },
      },
      disconnect() {},
    };

    await expect(probeClassic(port, "chrome-extension://test/")).rejects.toThrow("Specified native messaging host not found.");
    expect(lastErrorReads).toBe(1);
  });
});

describe("delivery", () => {
  it("writes the archive before the transport and does not create a second draft", async () => {
    const phases: string[] = [];
    let writes = 0;
    let creates = 0;
    const transport: ReportTransport = {
      async run(_draft, onPhase) {
        creates += 1;
        onPhase("opening");
        return { status: "created-not-opened", id: "m1", location: "https://outlook.office.com/mail/drafts", error: "tab blocked" };
      },
      async retryOpen() {
        return { status: "created-and-opened", location: "https://outlook.office.com/mail/drafts" };
      },
    };
    const outcome = await deliverReport({
      draft: draft(),
      client: "microsoft",
      gmailEnabled: true,
      archive: { async write() { writes += 1; return "2026-11-04 Daily Report.eml"; } },
      transport,
      onPhase: (phase) => phases.push(phase),
    });
    expect(writes).toBe(1);
    expect(creates).toBe(1);
    expect(phases).toEqual(["validating", "writing-archive", "creating", "opening", "partial"]);
    expect(outcome.canRetryOpen).toBe(true);
    rememberRetry(transport, outcome);
    const opened = await retryOpen();
    expect(creates).toBe(1);
    expect(opened.status).toBe("created-and-opened");
  });

  it("does not archive or open a second client when validation or the folder fails", async () => {
    let writes = 0;
    let runs = 0;
    const transport: ReportTransport = { async run() { runs += 1; return { status: "created-and-opened", location: "mailto" }; } };
    const blocked = await deliverReport({
      draft: draft(),
      client: "gmail",
      gmailEnabled: false,
      archive: { async write() { writes += 1; return "x.eml"; } },
      transport,
    });
    expect(blocked.status).toBe("no-side-effect-failure");
    expect(blocked.archived).toBe(false);
    const folder = await deliverReport({
      draft: draft(),
      client: "classic",
      archive: { async write() { writes += 1; throw new Error("denied"); } },
      transport,
    });
    expect(folder.message).toMatch(/Grant the job folder again/);
    expect(writes).toBe(1);
    expect(runs).toBe(0);
    const indeterminate: TransportResult = { status: "indeterminate", error: "timeout" };
    rememberRetry({ async run() { runs += 1; return indeterminate; } }, { status: "indeterminate", message: "", archived: true, canRetryOpen: false });
    const retry = await retryOpen();
    expect(retry.message).toMatch(/new draft was not created/i);
    expect(runs).toBe(0);
  });
});

describe("mailbox drafts", () => {
  it("stops after a Graph draft exists when opening fails", async () => {
    const urls: string[] = [];
    const result = await createMicrosoftDraft(draft(), "token", "crew@phxtech.com", {
      fetchImpl: async (url) => {
        urls.push(url);
        if (url.endsWith("/messages")) return { ok: true, status: 201, json: async () => ({ id: "m1" }), text: async () => "" };
        return { ok: true, status: 201, json: async () => ({}), text: async () => "" };
      },
      open: async () => { throw new Error("blocked"); },
      onOpening: () => undefined,
    });
    expect(result.status).toBe("created-not-opened");
    if (result.status === "created-not-opened") expect(result.location).toBe("https://outlook.office.com/mail/drafts");
    expect(urls).toHaveLength(2);
    const refused = await createMicrosoftDraft(draft(), "token", "a@outlook.com", {
      fetchImpl: async (url) => ({ ok: false, status: 400, json: async () => ({}), text: async () => url }),
      open: async () => undefined,
      onOpening: () => undefined,
    });
    expect(refused.status).toBe("no-side-effect-failure");
  });

  it("does not create a second Gmail draft when Drafts fails to open", async () => {
    let posts = 0;
    const result = await createGmailDraft(draft(), "token", {
      fetchImpl: async () => {
        posts += 1;
        return { ok: true, status: 200, json: async () => ({ id: "g1" }) };
      },
      open: async () => { throw new Error("blocked"); },
      onOpening: () => undefined,
    });
    expect(result.status).toBe("created-not-opened");
    expect(posts).toBe(1);
  });
});

describe("classic port and detection", () => {
  it("sends begin, chunks, and one commit", async () => {
    const script = scripted([
      { ok: true, command: "compose.begin" },
      { ok: true, command: "file.chunk" },
      { ok: true, command: "compose.commit", displayed: true },
    ]);
    const result = await composeOnPort(script.port, draft(), "chrome-extension://glkcmpmiomoppdkmdcjaepneimgbipdl/", () => undefined);
    expect(result.status).toBe("created-and-opened");
    const commands = script.sent.map((message) => (message as { command: string }).command);
    expect(commands).toEqual(["compose.begin", "file.chunk", "compose.commit"]);
    expect(script.sent.filter((message) => (message as { command: string }).command === "compose.commit")).toHaveLength(1);
  });

  it("suggests one client, defaults to outlook when none detected, and leaves a tie unselected", () => {
    expect(looksLikeClassic({ outlookRegistered: true, mailtoCommand: "outlook.exe", defaultClient: "" })).toBe(true);
    expect(looksLikeClassic({ outlookRegistered: true, mailtoCommand: "", defaultClient: "New Outlook" })).toBe(false);
    expect(chooseSuggestion({ classicReady: true, tabVendors: [], connected: null }).client).toBe("classic");
    expect(chooseSuggestion({ classicReady: true, tabVendors: ["gmail"], connected: null }).client).toBeNull();
    expect(chooseSuggestion({ classicReady: false, tabVendors: [], connected: "microsoft" }).client).toBe("microsoft");
    expect(chooseSuggestion({ classicReady: false, tabVendors: [], connected: null }).client).toBe("outlook");
  });

  it("opens an X-Unsent: 1 .eml draft directly in Classic or New Outlook via chrome.downloads.open", async () => {
    const listeners = new Set<(delta: DownloadChangeDelta) => void>();
    const openedIds: number[] = [];
    let capturedBlob: Blob | null = null;
    let savedFilename = "";
    const downloads: DownloadsApi = {
      async download(options) {
        savedFilename = options.filename;
        setTimeout(() => {
          for (const listener of listeners) listener({ id: 42, state: { current: "complete" } });
        }, 0);
        return 42;
      },
      async search() { return [{ id: 42, state: "in_progress" }]; },
      async open(id) { openedIds.push(id); },
      onChanged: {
        addListener(listener) { listeners.add(listener); },
        removeListener(listener) { listeners.delete(listener); },
      },
    };
    const outcome = await openDesktopDraft(draft(), {
      downloads,
      createObjectUrl(blob) { capturedBlob = blob; return "blob:test"; },
      revokeObjectUrl() {},
      now: () => new Date(2026, 10, 24),
    });
    expect(outcome.downloadId).toBe(42);
    expect(outcome.result).toEqual({ status: "created-and-opened", id: "42", location: "Outlook" });
    expect(openedIds).toEqual([42]);
    expect(savedFilename).toBe("PHX-Reports/2026-11-24 Daily Report.eml");
    expect(await capturedBlob!.text()).toMatch(/^X-Unsent: 1\r\nTo: a@example\.com\r\n/);
  });

  it("returns created-not-opened when Chrome requires a fresh click, and retries opening the same draft", async () => {
    let attempts = 0;
    const downloads: DownloadsApi = {
      async download() { return 7; },
      async search() { return [{ id: 7, state: "complete" }]; },
      async open() {
        attempts += 1;
        if (attempts === 1) throw new Error("User gesture required");
      },
      onChanged: { addListener() {}, removeListener() {} },
    };
    const outcome = await openDesktopDraft(draft(), {
      downloads,
      createObjectUrl: () => "blob:retry",
      revokeObjectUrl: () => undefined,
    });
    expect(outcome.downloadId).toBe(7);
    expect(outcome.result.status).toBe("created-not-opened");
    const retried = await openDownloadedDraft(7, downloads);
    expect(retried).toEqual({ status: "created-and-opened", id: "7", location: "Outlook" });
  });

  it("encodes mailto spaces as %20 instead of + so Outlook subjects are clean", () => {
    const url = mailtoUrl(draft());
    expect(url).toContain("subject=Daily%20Drilling%20Summary%20-%20Job");
    expect(url).not.toContain("Daily+Drilling+Summary");
  });
});

describe("session token", () => {
  it("stores the access token in session storage only", async () => {
    const session = memory();
    const local = memory();
    const token: SessionToken = { vendor: "microsoft", accessToken: "secret", expiresAt: 1, email: "a@phxtech.com" };
    await saveSessionToken(session, local, token);
    expect(await readSessionToken(session)).toEqual(token);
    expect(JSON.stringify(local.values)).not.toContain("secret");
    await clearSessionToken(session, local);
    expect(await readSessionToken(session)).toBeNull();
  });
});
