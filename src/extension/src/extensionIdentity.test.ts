import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import manifest from "../manifest.config";
import { MAIL_HOST_PERMISSIONS, OPTIONAL_WEBMAIL_ORIGINS, PREVIEW_EXTENSION_ID, PREVIEW_EXTENSION_KEY } from "./extensionIdentity";
import { PREVIEW_VERSION } from "./version";

function extensionId(key: string): string {
  const digest = createHash("sha256").update(Buffer.from(key, "base64")).digest();
  return [...digest.subarray(0, 16)]
    .map((byte) => byte.toString(16).padStart(2, "0"))
    .join("")
    .split("")
    .map((nibble) => String.fromCharCode(Number.parseInt(nibble, 16) + 97))
    .join("");
}

describe("preview identity", () => {
  it("derives the unpacked extension id from the public key", () => {
    expect(extensionId(PREVIEW_EXTENSION_KEY)).toBe(PREVIEW_EXTENSION_ID);
  });

  it("keeps the manifest permissions and the single preview version", () => {
    const source = readFileSync(new URL("../manifest.config.ts", import.meta.url), "utf8");
    const pkg = JSON.parse(readFileSync(new URL("../package.json", import.meta.url), "utf8")) as { version: string };
    const declared = manifest as unknown as {
      version: string;
      key: string;
      permissions: string[];
      host_permissions: string[];
      optional_host_permissions: string[];
    };
    expect(pkg.version).toBe(PREVIEW_VERSION);
    expect(source).toContain("version: PREVIEW_VERSION");
    expect(declared.version).toBe(PREVIEW_VERSION);
    expect(declared.key).toBe(PREVIEW_EXTENSION_KEY);
    expect(declared.permissions).toContain("downloads.open");
    expect(declared.permissions).toContain("nativeMessaging");
    expect(declared.permissions).not.toContain("tabs");
    for (const host of MAIL_HOST_PERMISSIONS) expect(declared.host_permissions).toContain(host);
    expect(declared.optional_host_permissions).toEqual([...OPTIONAL_WEBMAIL_ORIGINS]);
    const requiredHosts = new Set<string>(MAIL_HOST_PERMISSIONS);
    expect(OPTIONAL_WEBMAIL_ORIGINS.filter((host) => requiredHosts.has(host))).toEqual([]);
  });
});
