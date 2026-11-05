// The session the API issues after a vendor sign-in. Short lived; the refresh token renews it.
// Both are HMAC signed, so the API keeps no session table.

import { base64url, base64urlDecode, sign, verify } from "./crypto";

export interface Session {
  userId: string;
  email: string;
  /** Milliseconds since epoch. */
  exp: number;
}

const SESSION_SECONDS = 60 * 60;
const REFRESH_SECONDS = 30 * 24 * 60 * 60;

export async function issueSession(user: { userId: string; email: string }, secret: string, now = Date.now()): Promise<{ token: string; refresh: string }> {
  return {
    token: await pack({ userId: user.userId, email: user.email, exp: now + SESSION_SECONDS * 1000 }, secret),
    refresh: await pack({ userId: user.userId, email: user.email, exp: now + REFRESH_SECONDS * 1000 }, secret),
  };
}

/** The session a token carries, or null when it is unsigned, malformed, or expired. */
export async function readSession(token: string, secret: string, now = Date.now()): Promise<Session | null> {
  const dot = token.indexOf(".");
  if (dot < 0) return null;
  const body = token.slice(0, dot);
  const signature = token.slice(dot + 1);
  if (!(await verify(body, signature, secret))) return null;
  const parsed: unknown = JSON.parse(new TextDecoder().decode(base64urlDecode(body)));
  if (!isSession(parsed) || parsed.exp < now) return null;
  return parsed;
}

async function pack(session: Session, secret: string): Promise<string> {
  const body = base64url(new TextEncoder().encode(JSON.stringify(session)));
  return `${body}.${await sign(body, secret)}`;
}

function isSession(value: unknown): value is Session {
  if (typeof value !== "object" || value == null) return false;
  const record = value as Record<string, unknown>;
  return typeof record.userId === "string" && typeof record.email === "string" && typeof record.exp === "number";
}
