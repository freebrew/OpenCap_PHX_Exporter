import {
  buildAuthorizeUrl,
  codeChallenge,
  emailFromIdToken,
  parseTokenResponse,
  randomText,
  tokenRequestBody,
} from "@phx/mail";
import { mailRedirectUri } from "../../extensionIdentity";
import { GOOGLE_DRAFT_SCOPE, MAIL_STORAGE, MICROSOFT_DRAFT_SCOPE } from "./providerConfig";
import { parseMailClient, type MailClient } from "./clients";

export interface KeyValue {
  get(keys: string[]): Promise<Record<string, unknown>>;
  set(values: Record<string, unknown>): Promise<void>;
  remove(keys: string[]): Promise<void>;
}

export interface SessionToken {
  vendor: "microsoft" | "gmail";
  accessToken: string;
  expiresAt: number;
  email: string;
}

export interface ClientConfig {
  microsoftClientId: string;
  googleClientId: string;
  tenant: string;
  gmailEnabled: boolean;
}

export async function saveSessionToken(session: KeyValue, local: KeyValue, token: SessionToken): Promise<void> {
  await session.set({ [MAIL_STORAGE.sessionToken]: token });
  await local.remove([MAIL_STORAGE.sessionToken]);
}

export async function readSessionToken(session: KeyValue): Promise<SessionToken | null> {
  const stored = (await session.get([MAIL_STORAGE.sessionToken]))[MAIL_STORAGE.sessionToken];
  if (!stored || typeof stored !== "object") return null;
  const token = stored as Partial<SessionToken>;
  if ((token.vendor !== "microsoft" && token.vendor !== "gmail") || typeof token.accessToken !== "string" || typeof token.expiresAt !== "number") return null;
  return { vendor: token.vendor, accessToken: token.accessToken, expiresAt: token.expiresAt, email: typeof token.email === "string" ? token.email : "" };
}

export async function clearSessionToken(session: KeyValue, local: KeyValue): Promise<void> {
  await session.remove([MAIL_STORAGE.sessionToken]);
  await local.remove([MAIL_STORAGE.sessionToken]);
}

export function chromeLocal(): KeyValue {
  return {
    async get(keys) { return chrome.storage.local.get(keys); },
    async set(values) { await chrome.storage.local.set(values); },
    async remove(keys) { await chrome.storage.local.remove(keys); },
  };
}

export function chromeSession(): KeyValue {
  return {
    async get(keys) { return chrome.storage.session.get(keys); },
    async set(values) { await chrome.storage.session.set(values); },
    async remove(keys) { await chrome.storage.session.remove(keys); },
  };
}

export async function readMailClient(local: KeyValue = chromeLocal()): Promise<MailClient | null> {
  const stored = (await local.get([MAIL_STORAGE.client]))[MAIL_STORAGE.client];
  return parseMailClient(stored);
}

export async function saveMailClient(client: MailClient, local: KeyValue = chromeLocal()): Promise<void> {
  await local.set({ [MAIL_STORAGE.client]: client });
}

function textSetting(stored: Record<string, unknown>, key: string): string {
  const value = stored[key];
  return typeof value === "string" ? value : "";
}

export async function readClientConfig(local: KeyValue = chromeLocal()): Promise<ClientConfig> {
  const stored = await local.get([MAIL_STORAGE.microsoftClientId, MAIL_STORAGE.googleClientId, MAIL_STORAGE.tenant, MAIL_STORAGE.gmailEnabled]);
  return {
    microsoftClientId: textSetting(stored, MAIL_STORAGE.microsoftClientId),
    googleClientId: textSetting(stored, MAIL_STORAGE.googleClientId),
    tenant: textSetting(stored, MAIL_STORAGE.tenant) || "common",
    gmailEnabled: stored[MAIL_STORAGE.gmailEnabled] === true,
  };
}

export async function saveClientConfig(config: ClientConfig, local: KeyValue = chromeLocal()): Promise<void> {
  await local.set({
    [MAIL_STORAGE.microsoftClientId]: config.microsoftClientId.trim(),
    [MAIL_STORAGE.googleClientId]: config.googleClientId.trim(),
    [MAIL_STORAGE.tenant]: config.tenant.trim() || "common",
    [MAIL_STORAGE.gmailEnabled]: config.gmailEnabled,
  });
}

function vendorConfig(vendor: "microsoft" | "gmail", config: ClientConfig, extensionId: string): { authorize: string; token: string; clientId: string; redirect: string; scope: string } | null {
  const redirect = mailRedirectUri(extensionId);
  switch (vendor) {
    case "microsoft":
      if (config.microsoftClientId === "") return null;
      return {
        authorize: `https://login.microsoftonline.com/${encodeURIComponent(config.tenant || "common")}/oauth2/v2.0/authorize`,
        token: `https://login.microsoftonline.com/${encodeURIComponent(config.tenant || "common")}/oauth2/v2.0/token`,
        clientId: config.microsoftClientId,
        redirect,
        scope: MICROSOFT_DRAFT_SCOPE,
      };
    case "gmail":
      if (config.googleClientId === "") return null;
      return {
        authorize: "https://accounts.google.com/o/oauth2/v2/auth",
        token: "https://oauth2.googleapis.com/token",
        clientId: config.googleClientId,
        redirect,
        scope: GOOGLE_DRAFT_SCOPE,
      };
    default: {
      const never: never = vendor;
      return never;
    }
  }
}

function launch(url: string): Promise<string | "canceled"> {
  return new Promise((resolve, reject) => {
    chrome.identity.launchWebAuthFlow({ url, interactive: true }, (callback) => {
      const message = chrome.runtime.lastError?.message ?? "";
      if (callback) resolve(callback);
      else if (/cancel|denied|did not approve/i.test(message) || message === "") resolve("canceled");
      else reject(new Error(message));
    });
  });
}

/** Session access token only. A refresh token in the response is not stored. */
export async function authorizeMailbox(vendor: "microsoft" | "gmail", config: ClientConfig, session = chromeSession(), local = chromeLocal()): Promise<SessionToken | "canceled"> {
  const vendorSetup = vendorConfig(vendor, config, chrome.runtime.id);
  if (!vendorSetup) throw new Error(vendor === "gmail" ? "Add the Gmail client id in Mail settings." : "Add the Microsoft client id in Mail settings.");
  const verifier = randomText(48);
  const state = randomText(16);
  await session.set({ [MAIL_STORAGE.pendingAuth]: { verifier, state, vendor } });
  const url = buildAuthorizeUrl(vendorSetup, { state, challenge: await codeChallenge(verifier), prompt: "select_account" });
  const callback = await launch(url);
  await session.remove([MAIL_STORAGE.pendingAuth]);
  if (callback === "canceled") return "canceled";
  const returned = new URL(callback);
  if (returned.searchParams.get("state") !== state) throw new Error("Mailbox sign-in did not match this request.");
  const code = returned.searchParams.get("code") ?? "";
  if (code === "") throw new Error("Mailbox sign-in did not return a code.");
  const response = await fetch(vendorSetup.token, {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body: tokenRequestBody(vendorSetup.clientId, { grant_type: "authorization_code", code, redirect_uri: vendorSetup.redirect, code_verifier: verifier }),
  });
  if (!response.ok) throw new Error(`Mailbox sign-in failed (${response.status}).`);
  const parsed = parseTokenResponse(await response.json());
  if (!parsed) throw new Error("Mailbox sign-in did not return a token.");
  const token: SessionToken = { vendor, accessToken: parsed.accessToken, expiresAt: Date.now() + parsed.expiresIn * 1000, email: emailFromIdToken(parsed.idToken) };
  await saveSessionToken(session, local, token);
  return token;
}

export async function usableToken(vendor: "microsoft" | "gmail", config: ClientConfig): Promise<SessionToken | "canceled"> {
  const current = await readSessionToken(chromeSession());
  if (current && current.vendor === vendor && current.expiresAt - 60_000 > Date.now()) return current;
  await clearSessionToken(chromeSession(), chromeLocal());
  return authorizeMailbox(vendor, config);
}
