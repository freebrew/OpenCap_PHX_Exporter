import { nativeHostOrigin, OPTIONAL_WEBMAIL_ORIGINS, WEBMAIL_TAB_ORIGINS } from "../../extensionIdentity";
import { connectClassic, probeClassic } from "./classicHost";
import { chooseSuggestion, looksLikeClassic } from "./detect";
import type { MailClient } from "./clients";
import { chromeSession, readSessionToken } from "./platform";

export async function suggestClient(): Promise<{ client: MailClient | null; reason: string }> {
  let classicReady = false;
  try {
    const probe = await probeClassic(connectClassic(), nativeHostOrigin(chrome.runtime.id));
    classicReady = probe.ok === true && looksLikeClassic({
      outlookRegistered: probe.outlookRegistered === true,
      mailtoCommand: probe.mailtoCommand ?? "",
      defaultClient: probe.defaultClient ?? "",
    });
  } catch {
    classicReady = false;
  }
  const tabVendors: Array<"microsoft" | "gmail"> = [];
  try {
    const granted = await chrome.permissions.contains({ origins: [...OPTIONAL_WEBMAIL_ORIGINS] });
    if (granted) {
      const tabs = await chrome.tabs.query({ url: [...WEBMAIL_TAB_ORIGINS] });
      for (const tab of tabs) {
        const url = tab.url ?? "";
        if (url.includes("mail.google.com")) tabVendors.push("gmail");
        else if (url.includes("outlook.office.com") || url.includes("outlook.live.com")) tabVendors.push("microsoft");
      }
    }
  } catch {
    // Open-tab detection is optional. The picker still works without it.
  }
  const token = await readSessionToken(chromeSession());
  return chooseSuggestion({
    classicReady,
    tabVendors,
    connected: token && token.expiresAt > Date.now() ? token.vendor : null,
  });
}

export async function helperStatus(): Promise<string> {
  try {
    const probe = await probeClassic(connectClassic(), nativeHostOrigin(chrome.runtime.id));
    if (!probe.ok) return probe.error ?? "The Outlook helper did not answer.";
    if (probe.protocol !== 1) return "The Outlook helper is a different version. Extract the package and run repair-host.cmd.";
    if (probe.outlookRegistered !== true) return "The helper is installed. Classic Outlook is not registered on this PC.";
    return probe.outlookRunning === true ? "Classic Outlook is running." : "Classic Outlook is installed. The helper is ready.";
  } catch {
    return "The Outlook helper is not installed.";
  }
}

export async function downloadHelper(): Promise<void> {
  await chrome.downloads.download({
    url: chrome.runtime.getURL("native-host/phx-outlook-host.zip"),
    filename: "phx-outlook-host.zip",
    saveAs: true,
  });
}

export async function allowMailTabDetection(): Promise<boolean> {
  return chrome.permissions.request({ origins: [...OPTIONAL_WEBMAIL_ORIGINS] });
}
