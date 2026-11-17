import { defineManifest } from "@crxjs/vite-plugin";
import { MAIL_HOST_PERMISSIONS, OPTIONAL_WEBMAIL_ORIGINS, PREVIEW_EXTENSION_KEY } from "./src/extensionIdentity.ts";
import { PREVIEW_VERSION } from "./src/version.ts";

export default defineManifest({
  manifest_version: 3,
  name: "PHX Slide Sheet",
  version: PREVIEW_VERSION,
  key: PREVIEW_EXTENSION_KEY,
  description: "Slide Sheet preview. Excel stays the authority until the in-app record switch.",
  minimum_chrome_version: "122",
  action: {
    default_title: "Open PHX Slide Sheet preview",
  },
  background: {
    service_worker: "src/background.ts",
    type: "module",
  },
  // scripting, webRequest and *.phxtech.com are the OpenCap exporter's own (read-only) needs.
  // nativeMessaging is the Classic Outlook helper. identity is the mailbox sign-in.
  permissions: ["storage", "unlimitedStorage", "alarms", "downloads", "scripting", "activeTab", "webRequest", "identity", "nativeMessaging"],
  host_permissions: ["https://fieldcap-cdn.phxtech.com/*", "*://*.phxtech.com/*", ...MAIL_HOST_PERMISSIONS],
  optional_host_permissions: [...OPTIONAL_WEBMAIL_ORIGINS],
  web_accessible_resources: [
    {
      resources: ["injected-spy.js"],
      matches: ["*://*.phxtech.com/*"],
    },
  ],
  content_scripts: [
    {
      matches: ["https://fieldcap-cdn.phxtech.com/*"],
      js: ["src/fieldcap-content.ts"],
      run_at: "document_idle",
    },
  ],
  content_security_policy: {
    extension_pages: "script-src 'self'; object-src 'self'; worker-src 'self';",
  },
});
