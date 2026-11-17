// First-run email check. The phone and the web sheet ask for the address the PDFs go to; the API
// mails a 6-digit code and the person types it back. No sign-in: the code is the proof, and the
// address is not kept once it has been confirmed or the code has lapsed.

import { pairingCode } from "./crypto";
import { hashCode } from "./pairing";

export const CODE_MINUTES = 15;
export const MAX_TRIES = 5;
export const CODES_PER_ADDRESS_PER_HOUR = 5;
export const CODES_PER_IP_PER_HOUR = 20;

export type StartResult = { kind: "sent"; to: string; code: string } | { kind: "invalid" } | { kind: "limited" };
export type VerifyResult = "ok" | "wrong" | "expired" | "locked" | "invalid";

/** The address in its stored form, or null when it is not an address. */
export function normalEmail(raw: unknown): string | null {
  if (typeof raw !== "string") return null;
  const address = raw.trim().toLowerCase();
  if (address.length > 254 || !/^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]{2,}$/.test(address)) return null;
  return address;
}

/** Six digits, nothing else. */
export function normalCode(raw: unknown): string | null {
  if (typeof raw !== "string") return null;
  const code = raw.replace(/\s/g, "");
  return /^\d{6}$/.test(code) ? code : null;
}

/** Start a check: count it against the address and the caller, store the hash, return the code to mail. */
export async function startEmailCode(env: Env, rawEmail: unknown, ip: string): Promise<StartResult> {
  const address = normalEmail(rawEmail);
  if (!address) return { kind: "invalid" };
  const byAddress = env.PAIRING.getByName(`onboard:${address}`);
  const byCaller = env.PAIRING.getByName(`onboard-ip:${ip}`);
  if (!(await byAddress.allow(CODES_PER_ADDRESS_PER_HOUR)) || !(await byCaller.allow(CODES_PER_IP_PER_HOUR))) return { kind: "limited" };
  const now = Date.now();
  const code = pairingCode();
  await env.DB.batch([
    env.DB.prepare("DELETE FROM email_codes WHERE expires_at < ?").bind(now),
    env.DB.prepare(
      "INSERT INTO email_codes (email, code_hash, expires_at, tries) VALUES (?, ?, ?, 0) ON CONFLICT(email) DO UPDATE SET code_hash = excluded.code_hash, expires_at = excluded.expires_at, tries = 0",
    ).bind(address, await hashCode(code), now + CODE_MINUTES * 60_000),
  ]);
  return { kind: "sent", to: address, code };
}

/** Check a typed code. A right code deletes the row; MAX_TRIES wrong ones burn it. */
export async function verifyEmailCode(env: Env, rawEmail: unknown, rawCode: unknown): Promise<VerifyResult> {
  const address = normalEmail(rawEmail);
  const code = normalCode(rawCode);
  if (!address || !code) return "invalid";
  const row = await env.DB.prepare("SELECT code_hash, expires_at, tries FROM email_codes WHERE email = ?")
    .bind(address)
    .first<{ code_hash: string; expires_at: number; tries: number }>();
  if (!row || row.expires_at < Date.now()) {
    if (row) await env.DB.prepare("DELETE FROM email_codes WHERE email = ?").bind(address).run();
    return "expired";
  }
  if (row.tries >= MAX_TRIES) return "locked";
  if (row.code_hash === (await hashCode(code))) {
    await env.DB.prepare("DELETE FROM email_codes WHERE email = ?").bind(address).run();
    return "ok";
  }
  await env.DB.prepare("UPDATE email_codes SET tries = tries + 1 WHERE email = ?").bind(address).run();
  return row.tries + 1 >= MAX_TRIES ? "locked" : "wrong";
}

/** The mail the person reads. */
export function codeMessage(code: string): { subject: string; text: string; html: string } {
  return {
    subject: `StrapSheets code ${code}`,
    text: `Your StrapSheets code is ${code}. It works for ${CODE_MINUTES} minutes. If you did not ask for it, ignore this message.`,
    html: `<p>Your StrapSheets code is <strong style="font-size:20px;letter-spacing:2px">${code}</strong>.</p><p>It works for ${CODE_MINUTES} minutes. If you did not ask for it, ignore this message.</p>`,
  };
}
