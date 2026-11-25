import { defineManifest } from "@crxjs/vite-plugin";
import { MAIL_HOST_PERMISSIONS, OPTIONAL_WEBMAIL_ORIGINS, PREVIEW_EXTENSION_KEY } from "./src/extensionIdentity.ts";
import { PREVIEW_VERSION } from "./src/version.ts";

export default defineManifest({
  manifest_version: 3,
  name: "OpenCap Slidesheet",
  version: PREVIEW_VERSION,
  key: PREVIEW_EXTENSION_KEY,
  description: "Directional drilling slide sheet workbook in Chrome with FieldCap sync, proximity gauges, 3D well paths, and daily email reports.",
  minimum_chrome_version: "122",
  action: {
    default_title: "Open OpenCap Slidesheet",
    default_icon: {
      16: "icons/icon16.png",
      32: "icons/icon32.png",
      48: "icons/icon48.png",
    },
  },
  icons: {
    16: "icons/icon16.png",
    32: "icons/icon32.png",
    48: "icons/icon48.png",
    128: "icons/icon128.png",
  },
  background: {
    service_worker: "src/background.ts",
    type: "module",
  },
  // scripting, webRequest and *.phxtech.com are the OpenCap exporter's own (read-only) needs.
  // downloads.open launches the X-Unsent: 1 .eml draft directly in Classic or New Outlook.
  // nativeMessaging is the optional Classic Outlook COM helper. identity is mailbox OAuth sign-in.
  permissions: ["storage", "unlimitedStorage", "alarms", "downloads", "downloads.open", "scripting", "activeTab", "webRequest", "identity", "nativeMessaging"],
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
