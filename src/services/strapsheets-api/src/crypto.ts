// Tokens and codes. The signing key is a wrangler secret; nothing here is hardcoded.

/** HMAC-SHA256 of the text, base64url. */
export async function sign(text: string, secret: string): Promise<string> {
  const key = await crypto.subtle.importKey("raw", bytes(secret), { name: "HMAC", hash: "SHA-256" }, false, ["sign"]);
  const mac = await crypto.subtle.sign("HMAC", key, bytes(text));
  return base64url(new Uint8Array(mac));
}

/** True when the signature matches. The comparison does not stop at the first differing byte. */
export async function verify(text: string, signature: string, secret: string): Promise<boolean> {
  const expected = await sign(text, secret);
  const a = bytes(expected);
  const b = bytes(signature);
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i += 1) diff |= (a[i] ?? 0) ^ (b[i] ?? 0);
  return diff === 0;
}

/** A short random token, base64url. */
export function randomToken(bytesLength = 32): string {
  const buffer = new Uint8Array(bytesLength);
  crypto.getRandomValues(buffer);
  return base64url(buffer);
}

/** A 6-digit pairing code. */
export function pairingCode(): string {
  const buffer = new Uint32Array(1);
  crypto.getRandomValues(buffer);
  return String((buffer[0] ?? 0) % 1_000_000).padStart(6, "0");
}

export function base64url(data: Uint8Array): string {
  let text = "";
  for (const byte of data) text += String.fromCharCode(byte);
  return btoa(text).replaceAll("+", "-").replaceAll("/", "_").replaceAll("=", "");
}

export function bytes(text: string): Uint8Array<ArrayBuffer> {
  const encoded = new TextEncoder().encode(text);
  const out = new Uint8Array(new ArrayBuffer(encoded.length));
  out.set(encoded);
  return out;
}

export function base64urlDecode(text: string): Uint8Array<ArrayBuffer> {
  const padded = text.replaceAll("-", "+").replaceAll("_", "/") + "===".slice((text.length + 3) % 4);
  const binary = atob(padded);
  const out = new Uint8Array(new ArrayBuffer(binary.length));
  for (let i = 0; i < binary.length; i += 1) out[i] = binary.charCodeAt(i);
  return out;
}
