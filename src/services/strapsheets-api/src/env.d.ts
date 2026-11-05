// The bindings the Worker uses. `wrangler types` regenerates worker-configuration.d.ts from
// wrangler.jsonc; this lists the same bindings plus the secrets, which `wrangler secret put` holds
// and which therefore never appear in the generated file. The two are kept in step by the test that
// reads both.

/// <reference path="../worker-configuration.d.ts" />

interface Env {
  DB: D1Database;
  KV: KVNamespace;
  EMAIL: SendEmail;
  PAIRING: DurableObjectNamespace<import("./pairing").PairingLimiter>;
  APP_BASE_URL: string;
  MAIL_FROM: string;
  GOOGLE_CLIENT_ID: string;
  GOOGLE_CLIENT_SECRET: string;
  MICROSOFT_CLIENT_ID: string;
  MICROSOFT_CLIENT_SECRET: string;
  MICROSOFT_TENANT: string;
  SHEETS_CLIENT_EMAIL: string;
  SHEETS_PRIVATE_KEY: string;
  SHEETS_SPREADSHEET_ID: string;
  TOKEN_SIGNING_KEY: string;
}
