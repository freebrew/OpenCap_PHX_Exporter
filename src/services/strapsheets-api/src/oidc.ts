// Vendor sign-in. Google and Microsoft host the page; this only builds the authorize URL, checks
// the returned state, and reads the identity out of the ID token after verifying its signature
// against the vendor's published keys. The sign-in asks for openid, email and profile only.

export type Vendor = "google" | "microsoft";

export interface VendorIdentity {
  vendor: Vendor;
  sub: string;
  email: string;
}

interface VendorConfig {
  authorize: string;
  token: string;
  jwks: string;
  clientId: string;
  clientSecret: string;
}

const VENDORS = {
  google: {
    authorize: "https://accounts.google.com/o/oauth2/v2/auth",
    token: "https://oauth2.googleapis.com/token",
    jwks: "https://www.googleapis.com/oauth2/v3/certs",
  },
  microsoft: {
    jwksHost: "https://login.microsoftonline.com",
  },
} as const;

export function vendorConfig(vendor: Vendor, env: Env): VendorConfig {
  switch (vendor) {
    case "google":
      return { ...VENDORS.google, clientId: env.GOOGLE_CLIENT_ID, clientSecret: env.GOOGLE_CLIENT_SECRET };
    case "microsoft": {
      const tenant = env.MICROSOFT_TENANT || "common";
      return {
        authorize: `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/authorize`,
        token: `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`,
        jwks: `${VENDORS.microsoft.jwksHost}/${tenant}/discovery/v2.0/keys`,
        clientId: env.MICROSOFT_CLIENT_ID,
        clientSecret: env.MICROSOFT_CLIENT_SECRET ?? "",
      };
    }
    default: {
      const never: never = vendor;
      return never;
    }
  }
}

/** The authorize URL the browser is sent to. `state` and `challenge` are remembered server side. */
export function authorizeUrl(vendor: Vendor, env: Env, state: string, challenge: string): string {
  const config = vendorConfig(vendor, env);
  const params = new URLSearchParams({
    client_id: config.clientId,
    response_type: "code",
    redirect_uri: `${env.APP_BASE_URL}/auth/${vendor}/callback`,
    scope: "openid email profile",
    state,
    code_challenge: challenge,
    code_challenge_method: "S256",
  });
  return `${config.authorize}?${params.toString()}`;
}

/** Trade the code for tokens. Returns the ID token, still unverified. */
export async function exchangeCode(vendor: Vendor, env: Env, code: string, verifier: string): Promise<string> {
  const config = vendorConfig(vendor, env);
  const fields = new URLSearchParams({
    client_id: config.clientId,
    code,
    code_verifier: verifier,
    grant_type: "authorization_code",
    redirect_uri: `${env.APP_BASE_URL}/auth/${vendor}/callback`,
  });
  if (config.clientSecret) fields.set("client_secret", config.clientSecret);
  const response = await fetch(config.token, {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body: fields,
  });
  if (!response.ok) throw new Error(`${vendor} token exchange failed (${response.status}).`);
  const body = (await response.json()) as { id_token?: string };
  if (!body.id_token) throw new Error(`${vendor} returned no ID token.`);
  return body.id_token;
}

/** The identity an ID token carries, once its signature checks against the vendor's keys. */
export async function identityFromIdToken(vendor: Vendor, env: Env, idToken: string, now = Date.now(), expectedNonce = ""): Promise<VendorIdentity> {
  const config = vendorConfig(vendor, env);
  if (!config.clientId) throw new Error("Sign-in is not configured.");
  const [headerPart, payloadPart, signaturePart] = idToken.split(".");
  if (!headerPart || !payloadPart || !signaturePart) throw new Error("Malformed ID token.");
  const header = decodePart(headerPart) as { kid?: string; alg?: string };
  if (header.alg !== "RS256" || !header.kid) throw new Error("Unsupported ID token algorithm.");
  const payload = decodePart(payloadPart) as { sub?: string; email?: string; preferred_username?: string; aud?: string | string[]; iss?: string; exp?: number; nonce?: string };
  const jwks = vendor === "microsoft" ? microsoftJwks(payload.iss) : config.jwks;
  if (!jwks) throw new Error("ID token issuer is not recognised.");
  const key = await vendorKey(jwks, header.kid);
  const valid = await crypto.subtle.verify(
    { name: "RSASSA-PKCS1-v1_5" },
    key,
    base64urlToBytes(signaturePart),
    new TextEncoder().encode(`${headerPart}.${payloadPart}`),
  );
  if (!valid) throw new Error("ID token signature does not match.");
  const audiences = Array.isArray(payload.aud) ? payload.aud : [payload.aud];
  if (!audiences.includes(config.clientId)) throw new Error("ID token audience does not match.");
  if (typeof payload.exp === "number" && payload.exp * 1000 < now) throw new Error("ID token has expired.");
  if (expectedNonce && payload.nonce !== expectedNonce) throw new Error("ID token nonce does not match.");
  const email = emailOf(payload);
  if (!payload.sub || !email) throw new Error("ID token carries no identity.");
  return { vendor, sub: payload.sub, email };
}

/** Microsoft stamps the tenant on the issuer. Only that host's keys are fetched. */
function microsoftJwks(iss: string | undefined): string | null {
  const tenant = /^https:\/\/login\.microsoftonline\.com\/([0-9a-f-]{36})\/v2\.0$/i.exec(iss ?? "")?.[1];
  return tenant ? `https://login.microsoftonline.com/${tenant}/discovery/v2.0/keys` : null;
}

/** Work accounts put the address in preferred_username; some accounts also send email. */
function emailOf(payload: { email?: string; preferred_username?: string }): string {
  for (const value of [payload.email, payload.preferred_username]) {
    if (typeof value === "string" && value.includes("@")) return value.toLowerCase();
  }
  return "";
}

async function vendorKey(jwksUrl: string, kid: string): Promise<CryptoKey> {
  const response = await fetch(jwksUrl);
  if (!response.ok) throw new Error(`Could not read vendor keys (${response.status}).`);
  const jwks = (await response.json()) as { keys: Array<JsonWebKey & { kid?: string }> };
  const jwk = jwks.keys.find((key) => key.kid === kid);
  if (!jwk) throw new Error("Vendor published no key for this token.");
  return crypto.subtle.importKey("jwk", jwk, { name: "RSASSA-PKCS1-v1_5", hash: "SHA-256" }, false, ["verify"]);
}

function decodePart(part: string): unknown {
  return JSON.parse(new TextDecoder().decode(base64urlToBytes(part)));
}

function base64urlToBytes(text: string): Uint8Array<ArrayBuffer> {
  const padded = text.replaceAll("-", "+").replaceAll("_", "/") + "===".slice((text.length + 3) % 4);
  const binary = atob(padded);
  const out = new Uint8Array(new ArrayBuffer(binary.length));
  for (let i = 0; i < binary.length; i += 1) out[i] = binary.charCodeAt(i);
  return out;
}
