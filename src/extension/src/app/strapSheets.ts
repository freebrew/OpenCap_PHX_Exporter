// Publish a job's inventory and crew to StrapSheets. FieldCap is only read; the snapshot travels
// to the StrapSheets API under the Job Number, signed in through the same Google or Microsoft
// sign-in the phone uses. The FieldCap session never leaves this browser.

import { itemsFromCsv, strapCandidate, type InventoryItem } from "../../../packages/strapsheets-core/src/index";

const API_BASE_KEY = "phx_strapsheets_api_base";
const SESSION_KEY = "phx_strapsheets_session";
const DEFAULT_API_BASE = "http://localhost:8787";

export interface CrewMember {
  name: string;
  role: string;
  email: string;
}

export interface StrapSession {
  token: string;
  refresh: string;
  email: string;
}

interface TokenResponse {
  token?: string;
  refresh?: string;
  email?: string;
  error?: string;
}

async function apiBase(): Promise<string> {
  const stored = await chrome.storage.local.get(API_BASE_KEY);
  const base = stored[API_BASE_KEY];
  return typeof base === "string" && base !== "" ? base.replace(/\/$/, "") : DEFAULT_API_BASE;
}

/** The remembered sign-in, refreshed when its access token has lapsed. Null when nobody is signed in. */
export async function strapSession(): Promise<StrapSession | null> {
  const stored = await chrome.storage.local.get(SESSION_KEY);
  const session = stored[SESSION_KEY] as StrapSession | undefined;
  if (!session?.refresh) return null;
  const base = await apiBase();
  const response = await fetch(`${base}/auth/refresh`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ refresh: session.refresh }),
  });
  if (!response.ok) {
    await chrome.storage.local.remove(SESSION_KEY);
    return null;
  }
  const body = (await response.json()) as TokenResponse;
  if (!body.token || !body.refresh) return null;
  const next: StrapSession = { token: body.token, refresh: body.refresh, email: session.email };
  await chrome.storage.local.set({ [SESSION_KEY]: next });
  return next;
}

/**
 * Sign in through the API's vendor page. The browser opens it, the API redirects back to this
 * extension with the code, and the token exchange finishes here. The vendor only sees the API.
 */
export function signInToStrapSheets(vendor: "google" | "microsoft"): Promise<StrapSession> {
  return new Promise((resolve, reject) => {
    void apiBase().then((base) => {
      const redirect = chrome.identity.getRedirectURL("strapsheets");
      const url = `${base}/auth/${vendor}/start?redirect=${encodeURIComponent(redirect)}`;
      chrome.identity.launchWebAuthFlow({ url, interactive: true }, (callback) => {
        if (chrome.runtime.lastError || !callback) {
          reject(new Error(chrome.runtime.lastError?.message || "Sign-in was cancelled."));
          return;
        }
        const returned = new URL(callback);
        const token = returned.searchParams.get("token") ?? "";
        const refresh = returned.searchParams.get("refresh") ?? "";
        const email = returned.searchParams.get("email") ?? "";
        if (!token || !refresh) {
          reject(new Error("Sign-in did not return a session."));
          return;
        }
        const session: StrapSession = { token, refresh, email };
        void chrome.storage.local.set({ [SESSION_KEY]: session }).then(() => resolve(session));
      });
    });
  });
}

export interface PublishResult {
  jobNumber: string;
  items: number;
  /** Inventory rows the filter left out: MWD, tubulars, bits, collars, pipe. */
  filteredOut: number;
  crew: number;
}

/**
 * Send one job's inventory and crew to the API. The inventory CSV is the exporter's; only the
 * tools worth a strap sheet go (the Data tab's 3rd Party Tools rule for the Other family).
 */
export async function publishSnapshot(session: StrapSession, jobNumber: string, inventoryCsv: string, crew: CrewMember[]): Promise<PublishResult> {
  const base = await apiBase();
  const all = itemsFromCsv(inventoryCsv);
  const items: InventoryItem[] = all.filter(strapCandidate);
  const response = await fetch(`${base}/v1/jobs/${encodeURIComponent(jobNumber)}/inventory`, {
    method: "PUT",
    headers: { "content-type": "application/json", authorization: `Bearer ${session.token}` },
    body: JSON.stringify({ jobNumber, pulledAt: new Date().toISOString(), items, crew }),
  });
  if (!response.ok) {
    const body = (await response.json().catch(() => ({}))) as { error?: string };
    throw new Error(body.error || `Publish failed (${response.status}).`);
  }
  return { jobNumber, items: items.length, filteredOut: all.length - items.length, crew: crew.length };
}

/** The crew CSV (header row, then one person a row) into members. Rows with no name are dropped. */
export function crewFromCsv(csv: string): CrewMember[] {
  const lines = csv.replace(/^\uFEFF/, "").split(/\r\n|\r|\n/).filter((line) => line !== "");
  const headers = splitCsv(lines[0] ?? "").map((header) => header.trim().toLowerCase());
  const column = (name: string) => headers.indexOf(name);
  const nameCol = column("name");
  const roleCol = column("role");
  const emailCol = column("email");
  const members: CrewMember[] = [];
  for (const line of lines.slice(1)) {
    const cells = splitCsv(line);
    const name = (cells[nameCol] ?? "").trim();
    if (name === "") continue;
    members.push({ name, role: (cells[roleCol] ?? "").trim(), email: (cells[emailCol] ?? "").trim().toLowerCase() });
  }
  return members;
}

function splitCsv(line: string): string[] {
  const out: string[] = [];
  let token = "";
  let quoted = false;
  for (let i = 0; i < line.length; i += 1) {
    const ch = line[i];
    if (quoted) {
      if (ch === '"' && line[i + 1] === '"') {
        token += '"';
        i += 1;
      } else if (ch === '"') quoted = false;
      else token += ch;
    } else if (ch === '"') quoted = true;
    else if (ch === ",") {
      out.push(token);
      token = "";
    } else token += ch;
  }
  out.push(token);
  return out;
}
