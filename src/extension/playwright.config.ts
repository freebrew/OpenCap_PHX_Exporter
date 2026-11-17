import { defineConfig } from "@playwright/test";

export default defineConfig({
  testDir: "./e2e",
  timeout: 180_000,
  use: {
    baseURL: "http://127.0.0.1:4173",
    headless: true,
  },
  webServer: {
    command: "node ../../node_modules/vite/bin/vite.js --host 127.0.0.1 --port 4173 --strictPort",
    env: { PHX_PREVIEW: "1" },
    cwd: import.meta.dirname,
    url: "http://127.0.0.1:4173/app.html",
    reuseExistingServer: false,
    timeout: 120_000,
  },
});
