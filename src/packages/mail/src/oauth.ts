import { decodeBase64Url, encodeBase64Url } from "./bytes";

export interface OAuthToken {
  accessToken: string;
  refreshToken: string;
  expiresIn: number;
  idToken: string;
}

export interface AuthorizeConfig {
  authorize: string;
  clientId: string;
  redirect: string;
  scope: string;
}

export function randomText(bytes: number): string {
  const data = crypto.getRandomValues(new Uint8Array(bytes));
  return encodeBase64Url(data);
}

export async function codeChallenge(verifier: string): Promise<string> {
  const digest = new Uint8Array(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier)));
  return encodeBase64Url(digest);
}

export function buildAuthorizeUrl(config: AuthorizeConfig, extras: {
  state: string;
  challenge: string;
  loginHint?: string;
  googleOffline?: boolean;
  prompt?: string;
}): string {
  const url = new URL(config.authorize);
  url.searchParams.set("client_id", config.clientId);
  url.searchParams.set("redirect_uri", config.redirect);
  url.searchParams.set("response_type", "code");
  url.searchParams.set("scope", config.scope);
  url.searchParams.set("state", extras.state);
  url.searchParams.set("code_challenge", extras.challenge);
  url.searchParams.set("code_challenge_method", "S256");
  if (extras.loginHint) url.searchParams.set("login_hint", extras.loginHint);
  if (extras.googleOffline) {
    url.searchParams.set("access_type", "offline");
    url.searchParams.set("include_granted_scopes", "true");
  }
  if (extras.prompt) url.searchParams.set("prompt", extras.prompt);
  return url.toString();
}

export function parseTokenResponse(value: unknown): OAuthToken | null {
  if (!value || typeof value !== "object") return null;
  const body = value as Record<string, unknown>;
  if (typeof body.access_token !== "string" || body.access_token === "") return null;
  if (typeof body.expires_in !== "number" || !Number.isFinite(body.expires_in)) return null;
  return {
    accessToken: body.access_token,
    refreshToken: typeof body.refresh_token === "string" ? body.refresh_token : "",
    expiresIn: body.expires_in,
    idToken: typeof body.id_token === "string" ? body.id_token : "",
  };
}

export function emailFromIdToken(idToken: string): string {
  const payload = idToken.split(".")[1];
  if (!payload) return "";
  try {
    const json = JSON.parse(new TextDecoder().decode(decodeBase64Url(payload))) as { email?: unknown; preferred_username?: unknown };
    if (typeof json.email === "string" && json.email !== "") return json.email;
    return typeof json.preferred_username === "string" ? json.preferred_username : "";
  } catch {
    return "";
  }
}

export function tokenRequestBody(clientId: string, fields: Record<string, string>): URLSearchParams {
  return new URLSearchParams({ client_id: clientId, ...fields });
}
