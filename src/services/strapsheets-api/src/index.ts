// StrapSheets API. Sign-in through Google or Microsoft, email pairing, the inventory the
// extension relays, and the entries the phone posts. The sheet holds the strap rows; D1 holds
// the accounts and the snapshots.

import { Hono } from "hono";
import { cors } from "hono/cors";
import { cardFileName, DEFAULT_THREADS, ENTRY_FAMILIES, entryPayloadSchema, measuredSerials, measureRangesSchema, renderCards, type EntryPayload, type InventorySnapshot, type MeasureRange } from "../../../packages/strapsheets-core/src/index";
import { randomToken } from "./crypto";
import { authorizeUrl, exchangeCode, identityFromIdToken, type Vendor } from "./oidc";
import { codeMessage, normalEmail, startEmailCode, verifyEmailCode } from "./onboard";
import { PairingLimiter, pairingMessage, startPairing, verifyPairingCode, verifyPairingLink } from "./pairing";
import { issueSession, readSession } from "./session";
import { appendEntry, readFormOptions, sheetConfigured, writeRanges } from "./sheets";

export { PairingLimiter };

const HOURLY_SEND_CAP = 30;
/** The phone's deep link. The extension's chrome.identity redirect is any *.chromiumapp.org host. */
const APP_REDIRECT_SCHEME = "com.phx.strapsheets:";

const app = new Hono<{ Bindings: Env }>();

app.use("*", cors({
  origin: (origin) => allowedApiOrigin(origin) ? origin : null,
  allowMethods: ["GET", "POST", "PUT", "DELETE", "OPTIONS"],
  allowHeaders: ["Content-Type", "Authorization"],
}));

app.get("/health", (c) => c.json({ ok: true }));

app.get("/auth/:vendor/start", async (c) => {
  const vendor = vendorOf(c.req.param("vendor"));
  if (!vendor) return c.json({ error: "Unknown sign-in vendor." }, 404);
  const verifier = randomToken();
  const state = randomToken(16);
  const challenge = await pkceChallenge(verifier);
  // Where to send the session afterwards: the extension's chrome.identity redirect or the phone's
  // deep link. Anything else would hand the token to a stranger, so it is refused.
  const redirect = c.req.query("redirect") ?? "";
  if (redirect && !allowedRedirect(redirect)) return c.json({ error: "That sign-in redirect is not allowed." }, 400);
  await c.env.KV.put(`oidc:${state}`, JSON.stringify({ verifier, vendor, redirect }), { expirationTtl: 600 });
  return c.redirect(authorizeUrl(vendor, c.env, state, challenge));
});

app.get("/auth/:vendor/callback", async (c) => {
  const vendor = vendorOf(c.req.param("vendor"));
  const state = c.req.query("state") ?? "";
  const code = c.req.query("code") ?? "";
  if (!vendor || !code) return c.json({ error: "Sign-in did not complete." }, 400);
  const stored = await c.env.KV.get(`oidc:${state}`);
  if (!stored) return c.json({ error: "Sign-in state was not recognised." }, 400);
  const pending = JSON.parse(stored) as { verifier?: string; vendor?: string; redirect?: string };
  await c.env.KV.delete(`oidc:${state}`);
  if (pending.vendor !== vendor || !pending.verifier) return c.json({ error: "Sign-in state does not match." }, 400);
  const idToken = await exchangeCode(vendor, c.env, code, pending.verifier);
  const identity = await identityFromIdToken(vendor, c.env, idToken);
  const user = await upsertUser(c.env, identity.vendor, identity.sub, identity.email);
  const session = await issueSession(user, c.env.TOKEN_SIGNING_KEY);
  if (pending.redirect) {
    const back = new URL(pending.redirect);
    back.searchParams.set("token", session.token);
    back.searchParams.set("refresh", session.refresh);
    back.searchParams.set("email", user.email);
    return c.redirect(back.toString());
  }
  return c.json({ ...session, email: user.email });
});

app.post("/auth/refresh", async (c) => {
  const body = (await c.req.json()) as { refresh?: string };
  const session = await readSession(body.refresh ?? "", c.env.TOKEN_SIGNING_KEY);
  if (!session) return c.json({ error: "Refresh token is not valid." }, 401);
  return c.json(await issueSession(session, c.env.TOKEN_SIGNING_KEY));
});

app.post("/auth/pair/start", async (c) => {
  const session = await requireSession(c);
  if (session instanceof Response) return session;
  const body = (await c.req.json()) as { email?: string };
  if (!body.email) return c.json({ error: "An email address is required." }, 400);
  if (body.email.trim().toLowerCase() === session.email) {
    await c.env.DB.prepare("INSERT INTO pairings (email, user_id, verified_at) VALUES (?, ?, ?) ON CONFLICT(email) DO NOTHING")
      .bind(session.email, session.userId, new Date().toISOString())
      .run();
    return c.json({ paired: true, email: session.email });
  }
  const mail = await startPairing(c.env, session.userId, body.email);
  if (!mail) return c.json({ error: "Too many pairing codes for that address. Try again in an hour." }, 429);
  const message = pairingMessage(mail);
  await c.env.EMAIL.send({ to: mail.to, from: c.env.MAIL_FROM, replyTo: c.env.MAIL_FROM, ...message });
  return c.json({ paired: false, sent: true });
});

app.post("/auth/pair/verify", async (c) => {
  const session = await requireSession(c);
  if (session instanceof Response) return session;
  const body = (await c.req.json()) as { email?: string; code?: string };
  const ok = await verifyPairingCode(c.env, session.userId, body.email ?? "", body.code ?? "");
  return ok ? c.json({ paired: true }) : c.json({ error: "That code is not valid." }, 400);
});

// First-run email check for the phones and the web sheet. Anonymous on purpose: the person has no
// account yet, so the limits are per address and per caller, and a code is the only proof.
app.post("/v1/onboard/email", async (c) => {
  const body = (await c.req.json().catch(() => ({}))) as { email?: unknown };
  const started = await startEmailCode(c.env, body.email, c.req.header("cf-connecting-ip") ?? "unknown");
  switch (started.kind) {
    case "invalid":
      return c.json({ error: "Type a full email address." }, 400);
    case "limited":
      return c.json({ error: "Too many codes for that address. Try again in an hour." }, 429);
    case "sent":
      try {
        await c.env.EMAIL.send({ to: started.to, from: c.env.MAIL_FROM, replyTo: c.env.MAIL_FROM, ...codeMessage(started.code) });
      } catch (error) {
        // The mail service's own message names accounts and domains. Log it; tell the person something they can act on.
        console.error("onboard mail failed", error instanceof Error ? error.message : String(error));
        return c.json({ error: "We could not send the code just now. Try again in a few minutes." }, 502);
      }
      return c.json({ sent: true });
    default: {
      const never: never = started;
      return never;
    }
  }
});

app.post("/v1/onboard/verify", async (c) => {
  const body = (await c.req.json().catch(() => ({}))) as { email?: unknown; code?: unknown };
  const result = await verifyEmailCode(c.env, body.email, body.code);
  switch (result) {
    case "ok":
      return c.json({ verified: true, email: normalEmail(body.email) });
    case "invalid":
      return c.json({ error: "Type the 6 digits from the email.", reason: result }, 400);
    case "wrong":
      return c.json({ error: "That code is not right.", reason: result }, 400);
    case "locked":
      return c.json({ error: "Too many wrong codes. Ask for a new one.", reason: result }, 400);
    case "expired":
      return c.json({ error: "That code has expired. Ask for a new one.", reason: result }, 400);
    default: {
      const never: never = result;
      return never;
    }
  }
});

app.get("/auth/pair/verify", async (c) => {
  const email = await verifyPairingLink(c.env, c.req.query("token") ?? "");
  return email ? c.text(`Paired ${email}. You can return to the app.`) : c.text("This pairing link has expired.", 400);
});

app.get("/v1/jobs", async (c) => {
  const session = await requireSession(c);
  if (session instanceof Response) return session;
  const rows = await c.env.DB.prepare("SELECT job_number, pulled_at FROM snapshots ORDER BY job_number").all<{ job_number: string; pulled_at: string }>();
  return c.json({ jobs: rows.results });
});

app.get("/v1/jobs/:jobNumber/inventory", async (c) => {
  const session = await requireSession(c);
  if (session instanceof Response) return session;
  const row = await c.env.DB.prepare("SELECT items_json, pulled_at FROM snapshots WHERE job_number = ?")
    .bind(c.req.param("jobNumber"))
    .first<{ items_json: string; pulled_at: string }>();
  if (!row) return c.json({ jobNumber: c.req.param("jobNumber"), items: [], pulledAt: null });
  return c.json({ jobNumber: c.req.param("jobNumber"), items: JSON.parse(row.items_json), pulledAt: row.pulled_at });
});

app.put("/v1/jobs/:jobNumber/inventory", async (c) => {
  const session = await requireSession(c);
  if (session instanceof Response) return session;
  const snapshot = (await c.req.json()) as InventorySnapshot;
  await c.env.DB.prepare(
    "INSERT INTO snapshots (job_number, pulled_at, published_by, items_json, crew_json) VALUES (?, ?, ?, ?, ?) ON CONFLICT(job_number) DO UPDATE SET pulled_at = excluded.pulled_at, published_by = excluded.published_by, items_json = excluded.items_json, crew_json = excluded.crew_json",
  )
    .bind(c.req.param("jobNumber"), snapshot.pulledAt, session.userId, JSON.stringify(snapshot.items), JSON.stringify(snapshot))
    .run();
  return c.json({ stored: snapshot.items.length });
});

app.post("/v1/entries", async (c) => {
  const session = await requireSession(c);
  if (session instanceof Response) return session;
  const parsed = entryPayloadSchema.safeParse(await c.req.json());
  if (!parsed.success) return c.json({ error: "The entry is not valid." }, 400);
  const payload = { which: parsed.data.which, cells: cellsOf(parsed.data.cells) };
  const id = crypto.randomUUID();
  const jobNumber = c.req.query("jobNumber") ?? "";
  await c.env.DB.prepare("INSERT INTO entries (id, job_number, user_id, which_family, cells_json, created_at) VALUES (?, ?, ?, ?, ?, ?)")
    .bind(id, jobNumber, session.userId, payload.which, JSON.stringify(payload.cells), new Date().toISOString())
    .run();
  const range = await appendEntry(c.env, payload);
  return c.json({ id, range }, 201);
});

app.get("/v1/entries/:jobNumber", async (c) => {
  const session = await requireSession(c);
  if (session instanceof Response) return session;
  const rows = await c.env.DB.prepare("SELECT id, which_family, cells_json, created_at FROM entries WHERE job_number = ? AND user_id = ? ORDER BY created_at")
    .bind(c.req.param("jobNumber"), session.userId)
    .all<{ id: string; which_family: string; cells_json: string; created_at: string }>();
  const entries = rows.results.map((row) => ({ id: row.id, which: row.which_family, cells: JSON.parse(row.cells_json) as Record<string, string>, createdAt: row.created_at }));
  return c.json({ entries });
});

/** Every serial on the job that anyone has strapped, so the phones can show what is left. */
app.get("/v1/jobs/:jobNumber/measured", async (c) => {
  const session = await requireSession(c);
  if (session instanceof Response) return session;
  const rows = await c.env.DB.prepare("SELECT which_family, cells_json FROM entries WHERE job_number = ?")
    .bind(c.req.param("jobNumber"))
    .all<{ which_family: string; cells_json: string }>();
  const entries = rows.results.map((row) => ({ which: row.which_family, cells: JSON.parse(row.cells_json) as Record<number, string> }));
  return c.json({ serials: [...measuredSerials(ENTRY_FAMILIES, entries)] });
});

app.get("/v1/options", async (c) => {
  const session = await requireSession(c);
  if (session instanceof Response) return session;
  if (sheetConfigured(c.env)) {
    const options = await readFormOptions(c.env);
    return c.json({ sheet: true, threads: options.threads.length > 0 ? options.threads : DEFAULT_THREADS, recipients: options.recipients, ranges: options.ranges });
  }
  return c.json({ sheet: false, threads: DEFAULT_THREADS, recipients: await allowList(c.env), ranges: await storedRanges(c.env) });
});

app.put("/v1/ranges", async (c) => {
  const session = await requireSession(c);
  if (session instanceof Response) return session;
  const parsed = measureRangesSchema.safeParse(((await c.req.json()) as { ranges?: unknown }).ranges);
  if (!parsed.success) return c.json({ error: "The ranges are not valid." }, 400);
  const ranges: MeasureRange[] = parsed.data;
  await c.env.DB.batch([
    c.env.DB.prepare("DELETE FROM ranges"),
    ...ranges.map((range) => c.env.DB.prepare("INSERT INTO ranges (family, step_id, min, max) VALUES (?, ?, ?, ?)").bind(range.family, range.stepId, range.min, range.max)),
  ]);
  if (sheetConfigured(c.env)) await writeRanges(c.env, ranges);
  return c.json({ stored: ranges.length });
});

app.post("/v1/pdf", async (c) => {
  const session = await requireSession(c);
  if (session instanceof Response) return session;
  const entries = entriesOf(((await c.req.json()) as { entries?: unknown[] }).entries);
  if (!entries) return c.json({ error: "The entries are not valid." }, 400);
  const pdf = renderCards(entries);
  return new Response(pdf.buffer as ArrayBuffer, { headers: { "content-type": "application/pdf" } });
});

app.post("/v1/email", async (c) => {
  const session = await requireSession(c);
  if (session instanceof Response) return session;
  const body = (await c.req.json()) as { to?: string | string[]; entries?: unknown[]; combined?: boolean };
  const to = (Array.isArray(body.to) ? body.to : [body.to ?? ""]).map((email) => email.trim().toLowerCase()).filter(Boolean);
  if (to.length === 0) return c.json({ error: "Choose at least one recipient." }, 400);
  const allowed = sheetConfigured(c.env) ? (await readFormOptions(c.env)).recipients : await allowList(c.env);
  const refused = to.filter((email) => !allowed.includes(email));
  if (refused.length > 0) return c.json({ error: `${refused.join(", ")} not on the recipient list.` }, 403);
  const entries = entriesOf(body.entries);
  if (!entries || entries.length === 0) return c.json({ error: "Choose at least one entry." }, 400);
  const hourAgo = Date.now() - 3_600_000;
  const sent = await c.env.DB.prepare("SELECT COUNT(*) AS n FROM send_log WHERE sent_at > ?").bind(hourAgo).first<{ n: number }>();
  if ((sent?.n ?? 0) >= HOURLY_SEND_CAP) return c.json({ error: "The hourly send cap has been reached." }, 429);
  const files = body.combined === false
    ? entries.map((entry) => ({ name: cardFileName(entry), pdf: renderCards([entry]) }))
    : [{ name: entries.length > 1 ? `${entries.length} strap sheets.pdf` : cardFileName(entries[0] as EntryPayload), pdf: renderCards(entries) }];
  const result = await c.env.EMAIL.send({
    to,
    from: c.env.MAIL_FROM,
    replyTo: session.email,
    subject: entries.length > 1 ? `${entries.length} strap sheets` : `Strap sheet: ${entries[0]?.which ?? "entry"}`,
    text: `Sent by ${session.email}. The strap sheet${entries.length > 1 ? "s are" : " is"} attached.`,
    html: `<p>Sent by ${escapeHtml(session.email)}. The strap sheet${entries.length > 1 ? "s are" : " is"} attached.</p>`,
    attachments: files.map((file) => ({ content: base64(file.pdf), filename: file.name, type: "application/pdf", disposition: "attachment" as const })),
  });
  await c.env.DB.prepare("INSERT INTO send_log (sent_at) VALUES (?)").bind(Date.now()).run();
  return c.json({ messageId: result.messageId, files: files.length });
});

app.onError((err, c) => {
  console.error(JSON.stringify({ message: "request failed", error: err.message }));
  return c.json({ error: err.message }, 500);
});

export default app;

function vendorOf(value: string): Vendor | null {
  switch (value) {
    case "google":
    case "microsoft":
      return value;
    default:
      return null;
  }
}

async function requireSession(c: { env: Env; req: { header: (name: string) => string | undefined } }): Promise<{ userId: string; email: string } | Response> {
  const header = c.req.header("authorization") ?? "";
  const token = header.startsWith("Bearer ") ? header.slice(7) : "";
  const session = await readSession(token, c.env.TOKEN_SIGNING_KEY);
  return session ?? new Response(JSON.stringify({ error: "Sign in required." }), { status: 401, headers: { "content-type": "application/json" } });
}

async function upsertUser(env: Env, vendor: string, sub: string, email: string): Promise<{ userId: string; email: string }> {
  const existing = await env.DB.prepare("SELECT id FROM users WHERE vendor = ? AND vendor_sub = ?").bind(vendor, sub).first<{ id: string }>();
  if (existing) return { userId: existing.id, email };
  const userId = crypto.randomUUID();
  await env.DB.prepare("INSERT INTO users (id, email, vendor, vendor_sub, created_at) VALUES (?, ?, ?, ?, ?)").bind(userId, email, vendor, sub, new Date().toISOString()).run();
  return { userId, email };
}

async function pkceChallenge(verifier: string): Promise<string> {
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier));
  return btoa(String.fromCharCode(...new Uint8Array(digest))).replaceAll("+", "-").replaceAll("/", "_").replaceAll("=", "");
}

function allowedApiOrigin(origin: string): boolean {
  if (origin === "https://localhost" || origin === "http://localhost" || origin === "capacitor://localhost") return true;
  try {
    const url = new URL(origin);
    if (url.protocol === "chrome-extension:" && /^[a-p]{32}$/.test(url.hostname)) return true;
    // The web sheet is an Apps Script page, served from a hashed *.googleusercontent.com host.
    if (url.protocol === "https:" && url.hostname.endsWith("-script.googleusercontent.com")) return true;
    const local = url.hostname === "localhost" || url.hostname === "127.0.0.1";
    return local && (url.protocol === "http:" || url.protocol === "https:");
  } catch {
    return false;
  }
}

function allowedRedirect(value: string): boolean {
  try {
    const url = new URL(value);
    return url.protocol === APP_REDIRECT_SCHEME || (url.protocol === "https:" && url.hostname.endsWith(".chromiumapp.org"));
  } catch {
    return false;
  }
}

async function allowList(env: Env): Promise<string[]> {
  const rows = await env.DB.prepare("SELECT email FROM allow_list ORDER BY email").all<{ email: string }>();
  return rows.results.map((row) => row.email.toLowerCase());
}

async function storedRanges(env: Env): Promise<MeasureRange[]> {
  const rows = await env.DB.prepare("SELECT family, step_id, min, max FROM ranges ORDER BY family, step_id").all<{ family: string; step_id: string; min: number | null; max: number | null }>();
  return rows.results.map((row) => ({ family: row.family, stepId: row.step_id, min: row.min, max: row.max }));
}

function entriesOf(raw: unknown[] | undefined): EntryPayload[] | null {
  const out: EntryPayload[] = [];
  for (const entry of raw ?? []) {
    const parsed = entryPayloadSchema.safeParse(entry);
    if (!parsed.success) return null;
    out.push({ which: parsed.data.which, cells: cellsOf(parsed.data.cells) });
  }
  return out;
}

function escapeHtml(text: string): string {
  return text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;");
}

function cellsOf(cells: Record<string, string>): Record<number, string> {
  const out: Record<number, string> = {};
  for (const [col, value] of Object.entries(cells)) out[Number(col)] = value;
  return out;
}

function base64(data: Uint8Array): string {
  let text = "";
  for (const byte of data) text += String.fromCharCode(byte);
  return btoa(text);
}
