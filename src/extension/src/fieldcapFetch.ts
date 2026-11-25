/**
 * One FieldCap read for both browsers. Chrome sends the request from the extension's own worker.
 * Edge does not attach the FieldCap login cookies to a worker request, so there the same request
 * runs inside an open FieldCap tab (the OpenCap exporter does the same for its pulls). If that tab
 * is missing or answers with an error, the worker request is the fallback on either browser.
 */
export interface FieldCapReply {
  ok: boolean;
  status: number;
  text: string;
}

const FIELDCAP_TABS = ["*://*.phxtech.com/*"];

export function runningInEdge(): boolean {
  const brands = (navigator as Navigator & { userAgentData?: { brands?: Array<{ brand?: string }> } }).userAgentData?.brands;
  if (Array.isArray(brands) && brands.some((entry) => /edge/i.test(String(entry?.brand ?? "")))) return true;
  return /\bEdg\//.test(navigator.userAgent);
}

async function fieldCapTabId(): Promise<number> {
  const tabs = await chrome.tabs.query({ url: FIELDCAP_TABS });
  return tabs.find((tab) => tab.id != null && /fieldcap/i.test(tab.url ?? ""))?.id ?? 0;
}

async function readInTab(tabId: number, url: string, accept: string): Promise<FieldCapReply | null> {
  try {
    const [injected] = await chrome.scripting.executeScript({
      target: { tabId },
      world: "MAIN",
      args: [url, accept],
      func: async (target: string, header: string): Promise<FieldCapReply> => {
        try {
          const response = await fetch(target, { credentials: "include", headers: { Accept: header } });
          return { ok: response.ok, status: response.status, text: await response.text() };
        } catch {
          return { ok: false, status: 0, text: "" };
        }
      },
    });
    const reply = injected?.result as FieldCapReply | undefined;
    return reply && reply.status > 0 ? reply : null;
  } catch {
    // The tab went away, or the page would not run the script. The worker request is next.
    return null;
  }
}

export async function fieldCapGet(url: string, accept: string): Promise<FieldCapReply> {
  if (runningInEdge()) {
    const tabId = await fieldCapTabId();
    const inTab = tabId ? await readInTab(tabId, url, accept) : null;
    if (inTab?.ok) return inTab;
  }
  const response = await fetch(url, { credentials: "include", headers: { accept } });
  return { ok: response.ok, status: response.status, text: response.ok ? await response.text() : "" };
}
