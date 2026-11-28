export const PENDING_UPDATE_KEY = "phx_pending_update";
export const REOPEN_AFTER_UPDATE_KEY = "phx_reopen_after_update";
export const UPDATE_CHECK_ALARM = "phx-update-check";
export const UPDATE_CHECK_PERIOD_MINUTES = 30;
export const APPLY_UPDATE_MESSAGE = "phx-apply-update";

export function isNewerVersion(candidate: unknown, current: string): boolean {
  const text = typeof candidate === "string" ? candidate.trim() : "";
  if (!text) return false;
  const a = text.split(".").map((n) => parseInt(n, 10) || 0);
  const b = current.split(".").map((n) => parseInt(n, 10) || 0);
  const len = Math.max(a.length, b.length);
  for (let i = 0; i < len; i += 1) {
    const diff = (a[i] ?? 0) - (b[i] ?? 0);
    if (diff !== 0) return diff > 0;
  }
  return false;
}

export interface AutoUpdateApi {
  runtime: {
    getManifest(): { version: string; update_url?: string };
    getURL(path: string): string;
    reload(): void;
    requestUpdateCheck?: (
      callback: (status: string, details?: { version?: string }) => void,
    ) => void;
    lastError?: { message?: string } | undefined;
    onUpdateAvailable: { addListener(fn: (details: { version?: string }) => void): void };
    onInstalled: { addListener(fn: (details: { reason: string }) => void): void };
    onStartup?: { addListener(fn: () => void): void };
    onMessage: { addListener(fn: (message: { type?: string }, sender: unknown, sendResponse: (res: unknown) => void) => boolean | undefined): void };
  };
  alarms: {
    get(name: string, callback: (alarm?: chrome.alarms.Alarm) => void): void;
    create(name: string, info: { periodInMinutes: number }): void;
  } & { onAlarm: { addListener(fn: (alarm: { name: string }) => void): void } };
  storage: {
    local: {
      get(keys: string | string[]): Promise<Record<string, unknown>>;
      set(items: Record<string, unknown>): Promise<void>;
      remove(keys: string | string[]): Promise<void>;
    };
  };
  tabs: {
    query(info: { url: string }): Promise<Array<{ id?: number }>>;
    create(info: { url: string }): Promise<unknown>;
    onRemoved: { addListener(fn: (tabId: number) => void): void };
  };
  action: {
    setBadgeText(details: { text: string }): Promise<void> | void;
    setBadgeBackgroundColor(details: { color: string }): Promise<void> | void;
    setBadgeTextColor?: (details: { color: string }) => Promise<void> | void;
    setTitle(details: { title: string }): Promise<void> | void;
  };
}

/**
 * Automatically checks the Chrome Web Store server on startup and every 30 minutes,
 * installing downloaded updates automatically whenever no Slidesheet tab is open.
 * If a Slidesheet tab IS open, shows a badge + toolbar button and auto-installs as
 * soon as the last Slidesheet tab closes (or when the user clicks Update).
 */
export function startAutoUpdate(api: AutoUpdateApi = chrome as unknown as AutoUpdateApi): {
  checkNow(): void;
  tryInstall(): Promise<boolean>;
} {
  const manifest = api.runtime.getManifest();
  const currentVersion = manifest.version;
  const fromStore = typeof manifest.update_url === "string" && manifest.update_url.length > 0;
  const appUrl = api.runtime.getURL("app.html");

  async function setBadge(version: string | null): Promise<void> {
    try {
      if (version) {
        await api.action.setBadgeBackgroundColor({ color: "#1e6b35" });
        await api.action.setBadgeTextColor?.({ color: "#ffffff" });
        await api.action.setBadgeText({ text: "UP" });
        await api.action.setTitle({
          title: `OpenCap Slidesheet ${version} ready — installs automatically when tab closes`,
        });
      } else {
        await api.action.setBadgeText({ text: "" });
        await api.action.setTitle({ title: "Open OpenCap Slidesheet" });
      }
    } catch {
      // Ignore badge errors in headless/test environments.
    }
  }

  async function openAppTabCount(): Promise<number> {
    const tabs = await api.tabs.query({ url: `${appUrl}*` });
    return tabs.length;
  }

  async function tryInstall(): Promise<boolean> {
    const stored = await api.storage.local.get(PENDING_UPDATE_KEY);
    const pending = stored[PENDING_UPDATE_KEY];
    if (!isNewerVersion(pending, currentVersion)) {
      if (pending) await api.storage.local.remove(PENDING_UPDATE_KEY);
      await setBadge(null);
      return false;
    }
    const version = String(pending);
    if ((await openAppTabCount()) > 0) {
      await setBadge(version);
      return false;
    }
    await setBadge(null);
    await api.storage.local.remove(PENDING_UPDATE_KEY);
    api.runtime.reload();
    return true;
  }

  async function recordPending(version: string): Promise<void> {
    if (!isNewerVersion(version, currentVersion)) return;
    await api.storage.local.set({ [PENDING_UPDATE_KEY]: version });
    await tryInstall();
  }

  function scheduleCheckAlarm(): void {
    if (!fromStore) return;
    api.alarms.get(UPDATE_CHECK_ALARM, (existing) => {
      if (!existing) {
        api.alarms.create(UPDATE_CHECK_ALARM, { periodInMinutes: UPDATE_CHECK_PERIOD_MINUTES });
      }
    });
  }

  function checkNow(): void {
    if (!fromStore || typeof api.runtime.requestUpdateCheck !== "function") return;
    try {
      api.runtime.requestUpdateCheck((status, details) => {
        if (api.runtime.lastError) return;
        if (status === "update_available") {
          const version = details?.version ? String(details.version) : "";
          if (version) void recordPending(version);
          else void tryInstall();
        }
      });
    } catch {
      // Ignore transient Chrome Web Store check errors; the next 30-min alarm retries.
    }
  }

  api.runtime.onUpdateAvailable.addListener((details) => {
    const version = details?.version ? String(details.version) : "";
    if (version) {
      void recordPending(version);
    } else {
      void (async () => {
        if ((await openAppTabCount()) === 0) api.runtime.reload();
      })();
    }
  });

  api.tabs.onRemoved.addListener(() => {
    void tryInstall();
  });

  api.runtime.onMessage.addListener((message, _sender, sendResponse) => {
    if (message?.type !== APPLY_UPDATE_MESSAGE) return undefined;
    void (async () => {
      await api.storage.local.set({ [REOPEN_AFTER_UPDATE_KEY]: true });
      await api.storage.local.remove(PENDING_UPDATE_KEY);
      await setBadge(null);
      sendResponse({ ok: true });
      api.runtime.reload();
    })();
    return true;
  });

  api.runtime.onInstalled.addListener((details) => {
    void (async () => {
      await api.storage.local.remove(PENDING_UPDATE_KEY);
      await setBadge(null);
      scheduleCheckAlarm();
      if (details.reason === "update") {
        const stored = await api.storage.local.get(REOPEN_AFTER_UPDATE_KEY);
        if (stored[REOPEN_AFTER_UPDATE_KEY]) {
          await api.storage.local.remove(REOPEN_AFTER_UPDATE_KEY);
          if ((await openAppTabCount()) === 0) {
            await api.tabs.create({ url: appUrl });
          }
        }
      }
    })();
  });

  api.runtime.onStartup?.addListener(() => {
    scheduleCheckAlarm();
    checkNow();
  });

  api.alarms.onAlarm.addListener((alarm) => {
    if (alarm.name === UPDATE_CHECK_ALARM) {
      checkNow();
    }
  });

  scheduleCheckAlarm();
  void tryInstall();

  return { checkNow, tryInstall };
}
