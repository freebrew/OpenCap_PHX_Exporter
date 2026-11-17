const CHUNK = 0x8000;

export function encodeBase64(bytes: Uint8Array): string {
  const parts: string[] = [];
  for (let offset = 0; offset < bytes.length; offset += CHUNK) {
    parts.push(String.fromCharCode(...bytes.subarray(offset, offset + CHUNK)));
  }
  return btoa(parts.join(""));
}

export function decodeBase64(value: string): Uint8Array {
  const binary = atob(value.replace(/\s/g, ""));
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
  return bytes;
}

export function encodeBase64Url(bytes: Uint8Array): string {
  return encodeBase64(bytes).replaceAll("+", "-").replaceAll("/", "_").replaceAll("=", "");
}

export function decodeBase64Url(value: string): Uint8Array {
  const pad = value.length % 4 === 0 ? "" : "=".repeat(4 - (value.length % 4));
  return decodeBase64(value.replaceAll("-", "+").replaceAll("_", "/") + pad);
}

export function textBase64Url(text: string): string {
  return encodeBase64Url(new TextEncoder().encode(text));
}

export function wrapBase64(value: string): string {
  const lines: string[] = [];
  for (let index = 0; index < value.length; index += 76) lines.push(value.slice(index, index + 76));
  return lines.join("\r\n");
}
