import { fileURLToPath, URL } from "node:url";
import { defineConfig } from "vitest/config";

export default defineConfig({
  resolve: {
    alias: {
      "@phx/job-store": fileURLToPath(new URL("../packages/job-store/src/index.ts", import.meta.url)),
      "@phx/mail": fileURLToPath(new URL("../packages/mail/src/index.ts", import.meta.url)),
      "@phx/xlsx-bridge": fileURLToPath(new URL("../packages/xlsx-bridge/src/index.ts", import.meta.url)),
    },
  },
  test: {
    environment: "node",
    include: ["src/**/*.test.ts"],
  },
});
