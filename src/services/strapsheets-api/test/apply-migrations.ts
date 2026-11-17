import { env } from "cloudflare:workers";

// The schema from migrations/*.sql, inlined because the test runtime has no filesystem.
const statements = [
  `CREATE TABLE users (
    id TEXT PRIMARY KEY,
    email TEXT NOT NULL UNIQUE,
    vendor TEXT NOT NULL,
    vendor_sub TEXT NOT NULL,
    created_at TEXT NOT NULL
  )`,
  `CREATE TABLE pairings (
    email TEXT PRIMARY KEY,
    user_id TEXT NOT NULL REFERENCES users(id),
    verified_at TEXT NOT NULL
  )`,
  `CREATE TABLE pairing_codes (
    email TEXT PRIMARY KEY,
    code_hash TEXT NOT NULL,
    user_id TEXT NOT NULL REFERENCES users(id),
    expires_at INTEGER NOT NULL
  )`,
  `CREATE TABLE snapshots (
    job_number TEXT PRIMARY KEY,
    pulled_at TEXT NOT NULL,
    published_by TEXT NOT NULL,
    items_json TEXT NOT NULL,
    crew_json TEXT NOT NULL
  )`,
  `CREATE TABLE entries (
    id TEXT PRIMARY KEY,
    job_number TEXT NOT NULL,
    user_id TEXT NOT NULL,
    which_family TEXT NOT NULL,
    cells_json TEXT NOT NULL,
    created_at TEXT NOT NULL
  )`,
  "CREATE INDEX entries_by_job ON entries (job_number, created_at)",
  "CREATE TABLE allow_list (email TEXT PRIMARY KEY)",
  "CREATE TABLE send_log (id INTEGER PRIMARY KEY AUTOINCREMENT, sent_at INTEGER NOT NULL)",
  "CREATE TABLE ranges (family TEXT NOT NULL, step_id TEXT NOT NULL, min REAL, max REAL, PRIMARY KEY (family, step_id))",
  "CREATE TABLE email_codes (email TEXT PRIMARY KEY, code_hash TEXT NOT NULL, expires_at INTEGER NOT NULL, tries INTEGER NOT NULL DEFAULT 0)",
];

await env.DB.batch(statements.map((statement) => env.DB.prepare(statement)));
