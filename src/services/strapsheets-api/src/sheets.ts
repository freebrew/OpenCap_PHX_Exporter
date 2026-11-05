// The Tool Strap Sheet workbook. The API signs a service-account token with WebCrypto and appends
// one row to Form Responses 1, the same sheet apps-script/entry_api.js writes. The key is a
// wrangler secret and never touches the repo.

import type { EntryPayload, MeasureRange } from "../../../packages/strapsheets-core/src/index";

const SHEETS_SCOPE = "https://www.googleapis.com/auth/spreadsheets";
const TOKEN_URL = "https://oauth2.googleapis.com/token";
const FORM_OPTIONS = "FormOptions";
/** FormOptions columns: A thread types, B1:C1 form id, D recipients, F:I the Ranges block. */
const RANGES_BLOCK = `${FORM_OPTIONS}!F1:I500`;
const RANGES_HEADER = ["Ranges: Family", "Step", "Min", "Max"];

/** Whether this deployment has a Tool Strap Sheet workbook. A bought copy may not. */
export function sheetConfigured(env: Env): boolean {
  return env.SHEETS_SPREADSHEET_ID !== "" && env.SHEETS_CLIENT_EMAIL !== "";
}

export interface FormOptions {
  threads: string[];
  recipients: string[];
  ranges: MeasureRange[];
}

/** Thread types, the mail allow-list and the Ranges block, read in one batch. */
export async function readFormOptions(env: Env): Promise<FormOptions> {
  const token = await serviceAccountToken(env);
  const ranges = [`${FORM_OPTIONS}!A2:A50`, `${FORM_OPTIONS}!D2:D50`, RANGES_BLOCK].map((range) => `ranges=${encodeURIComponent(range)}`).join("&");
  const response = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${env.SHEETS_SPREADSHEET_ID}/values:batchGet?${ranges}`, {
    headers: { authorization: `Bearer ${token}` },
  });
  if (!response.ok) throw new Error(`FormOptions read failed (${response.status}).`);
  const body = (await response.json()) as { valueRanges?: { values?: string[][] }[] };
  const [threads, recipients, block] = (body.valueRanges ?? []).map((range) => range.values ?? []);
  return {
    threads: column(threads ?? []),
    recipients: column(recipients ?? []).map((email) => email.toLowerCase()),
    ranges: (block ?? []).slice(1).filter((row) => row[0] && row[1]).map((row) => ({
      family: String(row[0]),
      stepId: String(row[1]),
      min: numberOrNull(row[2]),
      max: numberOrNull(row[3]),
    })),
  };
}

/** Replace the Ranges block with these rows (header kept, old rows cleared). */
export async function writeRanges(env: Env, ranges: MeasureRange[]): Promise<void> {
  const token = await serviceAccountToken(env);
  const base = `https://sheets.googleapis.com/v4/spreadsheets/${env.SHEETS_SPREADSHEET_ID}/values`;
  const cleared = await fetch(`${base}/${encodeURIComponent(RANGES_BLOCK)}:clear`, { method: "POST", headers: { authorization: `Bearer ${token}` } });
  if (!cleared.ok) throw new Error(`Ranges clear failed (${cleared.status}).`);
  const values = [RANGES_HEADER, ...ranges.map((range) => [range.family, range.stepId, range.min ?? "", range.max ?? ""])];
  const written = await fetch(`${base}/${encodeURIComponent(`${FORM_OPTIONS}!F1`)}?valueInputOption=RAW`, {
    method: "PUT",
    headers: { authorization: `Bearer ${token}`, "content-type": "application/json" },
    body: JSON.stringify({ values }),
  });
  if (!written.ok) throw new Error(`Ranges write failed (${written.status}).`);
}

function column(rows: string[][]): string[] {
  const out: string[] = [];
  for (const row of rows) {
    const value = String(row[0] ?? "").trim();
    if (value === "") break;
    out.push(value);
  }
  return out;
}

function numberOrNull(value: unknown): number | null {
  const text = String(value ?? "").trim();
  return text === "" || Number.isNaN(Number(text)) ? null : Number(text);
}

/** Append one entry as a row. Returns the sheet's updated range, or "" with no sheet. */
export async function appendEntry(env: Env, payload: EntryPayload): Promise<string> {
  if (!sheetConfigured(env)) return "";
  const token = await serviceAccountToken(env);
  const width = Math.max(0, ...Object.keys(payload.cells).map(Number)) + 1;
  const row = Array.from({ length: width }, (_, col) => payload.cells[col] ?? "");
  row[0] = new Date().toISOString();
  row[1] = payload.which;
  const response = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${env.SHEETS_SPREADSHEET_ID}/values/${encodeURIComponent("Form Responses 1")}!A1:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,
    {
      method: "POST",
      headers: { authorization: `Bearer ${token}`, "content-type": "application/json" },
      body: JSON.stringify({ values: [row] }),
    },
  );
  if (!response.ok) throw new Error(`Sheets append failed (${response.status}).`);
  const body = (await response.json()) as { updates?: { updatedRange?: string } };
  return body.updates?.updatedRange ?? "";
}

/** A one-hour access token for the service account, signed here and not stored. */
export async function serviceAccountToken(env: Env, now = Date.now()): Promise<string> {
  const header = base64url(JSON.stringify({ alg: "RS256", typ: "JWT" }));
  const claims = base64url(JSON.stringify({
    iss: env.SHEETS_CLIENT_EMAIL,
    scope: SHEETS_SCOPE,
    aud: TOKEN_URL,
    iat: Math.floor(now / 1000),
    exp: Math.floor(now / 1000) + 3600,
  }));
  const key = await crypto.subtle.importKey(
    "pkcs8",
    pemToDer(env.SHEETS_PRIVATE_KEY),
    { name: "RSASSA-PKCS1-v1_5", hash: "SHA-256" },
    false,
    ["sign"],
  );
  const signature = await crypto.subtle.sign("RSASSA-PKCS1-v1_5", key, new TextEncoder().encode(`${header}.${claims}`));
  const jwt = `${header}.${claims}.${base64urlBytes(new Uint8Array(signature))}`;
  const response = await fetch(TOKEN_URL, {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer", assertion: jwt }),
  });
  if (!response.ok) throw new Error(`Service account token failed (${response.status}).`);
  const body = (await response.json()) as { access_token?: string };
  if (!body.access_token) throw new Error("Service account returned no token.");
  return body.access_token;
}

function pemToDer(pem: string): ArrayBuffer {
  const body = pem.replace(/-----[^-]+-----/g, "").replace(/\s+/g, "");
  const binary = atob(body);
  const out = new Uint8Array(new ArrayBuffer(binary.length));
  for (let i = 0; i < binary.length; i += 1) out[i] = binary.charCodeAt(i);
  return out.buffer;
}

function base64url(text: string): string {
  return base64urlBytes(new TextEncoder().encode(text));
}

function base64urlBytes(data: Uint8Array): string {
  let text = "";
  for (const byte of data) text += String.fromCharCode(byte);
  return btoa(text).replaceAll("+", "-").replaceAll("/", "_").replaceAll("=", "");
}
