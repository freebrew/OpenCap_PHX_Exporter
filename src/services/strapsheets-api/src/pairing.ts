// Email pairing. The user types the email their crew sheet carries; the API mails a 6-digit code
// and a link. A signed-in email that already matches a crew email is paired with no code.

import { DurableObject } from "cloudflare:workers";
import { pairingCode, randomToken } from "./crypto";

const CODE_MINUTES = 15;
const MAX_CODES_PER_HOUR = 5;

/**
 * One limiter per email address. It counts how many codes that address has been sent in the last
 * hour, so a script cannot flood an inbox.
 */
export class PairingLimiter extends DurableObject<Env> {
  async allow(): Promise<boolean> {
    const hour = Math.floor(Date.now() / 3_600_000);
    const count = (await this.ctx.storage.get<number>(`h${hour}`)) ?? 0;
    if (count >= MAX_CODES_PER_HOUR) return false;
    await this.ctx.storage.put(`h${hour}`, count + 1);
    return true;
  }
}

export interface PairingMail {
  to: string;
  code: string;
  link: string;
}

/** Start a pairing. Returns the mail to send, or null when the address is being asked too often. */
export async function startPairing(env: Env, userId: string, email: string): Promise<PairingMail | null> {
  const address = email.trim().toLowerCase();
  const limiter = env.PAIRING.getByName(address);
  if (!(await limiter.allow())) return null;
  const code = pairingCode();
  await env.DB.prepare(
    "INSERT INTO pairing_codes (email, code_hash, user_id, expires_at) VALUES (?, ?, ?, ?) ON CONFLICT(email) DO UPDATE SET code_hash = excluded.code_hash, user_id = excluded.user_id, expires_at = excluded.expires_at",
  )
    .bind(address, await hashCode(code), userId, Date.now() + CODE_MINUTES * 60_000)
    .run();
  const token = randomToken();
  await env.KV.put(`pair:${token}`, JSON.stringify({ email: address, userId }), { expirationTtl: CODE_MINUTES * 60 });
  return { to: address, code, link: `${env.APP_BASE_URL}/auth/pair/verify?token=${token}` };
}

/** Confirm a pairing from the typed code. True when the code matches and has not expired. */
export async function verifyPairingCode(env: Env, userId: string, email: string, code: string): Promise<boolean> {
  const address = email.trim().toLowerCase();
  const row = await env.DB.prepare("SELECT code_hash, user_id, expires_at FROM pairing_codes WHERE email = ?")
    .bind(address)
    .first<{ code_hash: string; user_id: string; expires_at: number }>();
  if (!row || row.user_id !== userId || row.expires_at < Date.now()) return false;
  if (row.code_hash !== (await hashCode(code.trim()))) return false;
  await confirmPairing(env, userId, address);
  return true;
}

/** Confirm a pairing from the mailed link. Returns the paired email, or null when the link is stale. */
export async function verifyPairingLink(env: Env, token: string): Promise<string | null> {
  const stored = await env.KV.get(`pair:${token}`);
  if (!stored) return null;
  const parsed = JSON.parse(stored) as { email?: string; userId?: string };
  if (!parsed.email || !parsed.userId) return null;
  await confirmPairing(env, parsed.userId, parsed.email);
  await env.KV.delete(`pair:${token}`);
  return parsed.email;
}

async function confirmPairing(env: Env, userId: string, email: string): Promise<void> {
  await env.DB.batch([
    env.DB.prepare("INSERT INTO pairings (email, user_id, verified_at) VALUES (?, ?, ?) ON CONFLICT(email) DO UPDATE SET user_id = excluded.user_id, verified_at = excluded.verified_at").bind(email, userId, new Date().toISOString()),
    env.DB.prepare("DELETE FROM pairing_codes WHERE email = ?").bind(email),
  ]);
}

async function hashCode(code: string): Promise<string> {
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(code));
  return [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, "0")).join("");
}

/** The mail the user reads. Both the code and the link pair the address. */
export function pairingMessage(mail: PairingMail): { subject: string; text: string; html: string } {
  return {
    subject: "Your StrapSheets pairing code",
    text: `Your pairing code is ${mail.code}. It expires in ${CODE_MINUTES} minutes. Or open ${mail.link}`,
    html: `<p>Your pairing code is <strong>${mail.code}</strong>. It expires in ${CODE_MINUTES} minutes.</p><p><a href="${mail.link}">Pair this email</a></p>`,
  };
}
