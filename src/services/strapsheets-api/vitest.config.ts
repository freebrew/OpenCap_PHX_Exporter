import { cloudflareTest } from "@cloudflare/vitest-pool-workers";
import { defineConfig } from "vitest/config";

export default defineConfig({
  plugins: [
    cloudflareTest({
      wrangler: { configPath: "./wrangler.jsonc" },
      miniflare: { bindings: { TOKEN_SIGNING_KEY: "test-signing-key" } },
    }),
  ],
  test: {
    setupFiles: ["./test/apply-migrations.ts"],
  },
});
