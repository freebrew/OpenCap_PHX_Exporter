import { describe, expect, it, vi } from "vitest";
import {
  APPLY_UPDATE_MESSAGE,
  isNewerVersion,
  PENDING_UPDATE_KEY,
  REOPEN_AFTER_UPDATE_KEY,
  startAutoUpdate,
  UPDATE_CHECK_ALARM,
  UPDATE_CHECK_PERIOD_MINUTES,
  type AutoUpdateApi,
} from "./autoUpdate";

function makeFakeApi(opts: {
  version?: string;
  updateUrl?: string;
  openTabs?: number;
  checkStatus?: string;
  checkVersion?: string;
}) {
  const store = new Map<string, unknown>();
  const alarms = new Map<string, { periodInMinutes: number }>();
  let updateAvailableFn: ((d: { version?: string }) => void) | null = null;
  let installedFn: ((d: { reason: string }) => void) | null = null;
  let startupFn: (() => void) | null = null;
  let alarmFn: ((a: { name: string }) => void) | null = null;
  let tabRemovedFn: ((id: number) => void) | null = null;
  let messageFn: ((m: { type?: string }, s: unknown, r: (v: unknown) => void) => boolean | undefined) | null = null;
  let openTabs = opts.openTabs ?? 0;

  const reload = vi.fn();
  const createTab = vi.fn(async () => ({}));
  const setBadgeText = vi.fn();
  const requestUpdateCheck = vi.fn((cb: (status: string, details?: { version?: string }) => void) => {
    cb(opts.checkStatus ?? "no_update", opts.checkVersion ? { version: opts.checkVersion } : undefined);
  });

  const api: AutoUpdateApi = {
    runtime: {
      getManifest: () => ({
        version: opts.version ?? "0.1.104",
        ...(opts.updateUrl ? { update_url: opts.updateUrl } : {}),
      }),
      getURL: (p) => `chrome-extension://test/${p}`,
      reload,
      requestUpdateCheck,
      onUpdateAvailable: { addListener: (fn) => { updateAvailableFn = fn; } },
      onInstalled: { addListener: (fn) => { installedFn = fn; } },
      onStartup: { addListener: (fn) => { startupFn = fn; } },
      onMessage: { addListener: (fn) => { messageFn = fn; } },
    },
    alarms: {
      get: (name, cb) => cb(alarms.has(name) ? ({ name } as chrome.alarms.Alarm) : undefined),
      create: (name, info) => { alarms.set(name, info); },
      onAlarm: { addListener: (fn) => { alarmFn = fn; } },
    },
    storage: {
      local: {
        get: async (keys) => {
          const list = Array.isArray(keys) ? keys : [keys];
          const out: Record<string, unknown> = {};
          for (const k of list) if (store.has(k)) out[k] = store.get(k);
          return out;
        },
        set: async (items) => { for (const [k, v] of Object.entries(items)) store.set(k, v); },
        remove: async (keys) => {
          for (const k of Array.isArray(keys) ? keys : [keys]) store.delete(k);
        },
      },
    },
    tabs: {
      query: async () => Array.from({ length: openTabs }, (_, i) => ({ id: i + 1 })),
      create: createTab,
      onRemoved: { addListener: (fn) => { tabRemovedFn = fn; } },
    },
    action: {
      setBadgeText,
      setBadgeBackgroundColor: vi.fn(),
      setBadgeTextColor: vi.fn(),
      setTitle: vi.fn(),
    },
  };

  return {
    api,
    store,
    alarms,
    reload,
    createTab,
    setBadgeText,
    requestUpdateCheck,
    setOpenTabs: (n: number) => { openTabs = n; },
    fireUpdateAvailable: (version: string) => updateAvailableFn?.({ version }),
    fireInstalled: (reason: string) => installedFn?.({ reason }),
    fireStartup: () => startupFn?.(),
    fireAlarm: (name: string) => alarmFn?.({ name }),
    fireTabRemoved: () => tabRemovedFn?.(1),
    fireMessage: (msg: { type?: string }) => new Promise((resolve) => { messageFn?.(msg, {}, resolve); }),
  };
}

const flush = () => new Promise((r) => setTimeout(r, 15));

describe("autoUpdate", () => {
  it("compares dotted versions accurately", () => {
    expect(isNewerVersion("0.1.105", "0.1.104")).toBe(true);
    expect(isNewerVersion("0.2.0", "0.1.104")).toBe(true);
    expect(isNewerVersion("0.1.104", "0.1.104")).toBe(false);
    expect(isNewerVersion("0.1.99", "0.1.104")).toBe(false);
    expect(isNewerVersion("", "0.1.104")).toBe(false);
  });

  it("does not schedule store checks when loaded unpacked (no update_url)", () => {
    const fake = makeFakeApi({ version: "0.1.104" });
    startAutoUpdate(fake.api);
    expect(fake.alarms.has(UPDATE_CHECK_ALARM)).toBe(false);
    fake.fireStartup();
    expect(fake.requestUpdateCheck).not.toHaveBeenCalled();
  });

  it("schedules 30-minute checks and auto-reloads immediately when no tab is open", async () => {
    const fake = makeFakeApi({
      version: "0.1.104",
      updateUrl: "https://clients2.google.com/service/update2/crx",
      openTabs: 0,
      checkStatus: "update_available",
      checkVersion: "0.1.105",
    });
    startAutoUpdate(fake.api);
    expect(fake.alarms.get(UPDATE_CHECK_ALARM)).toEqual({ periodInMinutes: UPDATE_CHECK_PERIOD_MINUTES });

    fake.fireAlarm(UPDATE_CHECK_ALARM);
    await flush();
    expect(fake.requestUpdateCheck).toHaveBeenCalledTimes(1);
    expect(fake.reload).toHaveBeenCalledTimes(1);
    expect(fake.store.has(PENDING_UPDATE_KEY)).toBe(false);
  });

  it("defers reload while Slidesheet tab is open and auto-installs when the tab closes", async () => {
    const fake = makeFakeApi({
      version: "0.1.104",
      updateUrl: "https://clients2.google.com/service/update2/crx",
      openTabs: 1,
    });
    startAutoUpdate(fake.api);

    fake.fireUpdateAvailable("0.1.105");
    await flush();
    expect(fake.reload).not.toHaveBeenCalled();
    expect(fake.store.get(PENDING_UPDATE_KEY)).toBe("0.1.105");
    expect(fake.setBadgeText).toHaveBeenCalledWith({ text: "UP" });

    fake.setOpenTabs(0);
    fake.fireTabRemoved();
    await flush();
    expect(fake.reload).toHaveBeenCalledTimes(1);
    expect(fake.store.has(PENDING_UPDATE_KEY)).toBe(false);
  });

  it("reopens Slidesheet automatically after an explicit user-triggered update", async () => {
    const fake = makeFakeApi({
      version: "0.1.104",
      updateUrl: "https://clients2.google.com/service/update2/crx",
      openTabs: 1,
    });
    startAutoUpdate(fake.api);
    fake.store.set(PENDING_UPDATE_KEY, "0.1.105");

    await fake.fireMessage({ type: APPLY_UPDATE_MESSAGE });
    expect(fake.store.get(REOPEN_AFTER_UPDATE_KEY)).toBe(true);
    expect(fake.reload).toHaveBeenCalledTimes(1);

    fake.setOpenTabs(0);
    fake.fireInstalled("update");
    await flush();
    expect(fake.createTab).toHaveBeenCalledWith({ url: "chrome-extension://test/app.html" });
    expect(fake.store.has(REOPEN_AFTER_UPDATE_KEY)).toBe(false);
  });
});
