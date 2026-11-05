-- StrapSheets account data. The sheet stays the record of the strap entries; D1 holds the
-- accounts, the pairings, and the inventory the extension relays.

CREATE TABLE users (
  id TEXT PRIMARY KEY,
  email TEXT NOT NULL UNIQUE,
  vendor TEXT NOT NULL,
  vendor_sub TEXT NOT NULL,
  created_at TEXT NOT NULL
);

CREATE TABLE pairings (
  email TEXT PRIMARY KEY,
  user_id TEXT NOT NULL REFERENCES users(id),
  verified_at TEXT NOT NULL
);

-- A pairing code waiting to be typed. One live code per email; the newest replaces the last.
CREATE TABLE pairing_codes (
  email TEXT PRIMARY KEY,
  code_hash TEXT NOT NULL,
  user_id TEXT NOT NULL REFERENCES users(id),
  expires_at INTEGER NOT NULL
);

-- The inventory the extension published for one Job Number. One row per job, replaced wholesale.
CREATE TABLE snapshots (
  job_number TEXT PRIMARY KEY,
  pulled_at TEXT NOT NULL,
  published_by TEXT NOT NULL,
  items_json TEXT NOT NULL,
  crew_json TEXT NOT NULL
);

-- Entries the phone posted while offline, once they have synced. The sheet is the full record.
CREATE TABLE entries (
  id TEXT PRIMARY KEY,
  job_number TEXT NOT NULL,
  user_id TEXT NOT NULL,
  which_family TEXT NOT NULL,
  cells_json TEXT NOT NULL,
  created_at TEXT NOT NULL
);

CREATE INDEX entries_by_job ON entries (job_number, created_at);

-- Recipients the service mailbox will send to, and how many it has sent this hour.
CREATE TABLE allow_list (
  email TEXT PRIMARY KEY
);

CREATE TABLE send_log (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  sent_at INTEGER NOT NULL
);
