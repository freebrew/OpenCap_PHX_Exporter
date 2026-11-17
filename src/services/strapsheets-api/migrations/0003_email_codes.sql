-- Confirmation codes for the first-run email check on the phones and the web sheet. No account is
-- involved: a row lives from the moment a code is mailed until it is confirmed, burned by wrong
-- tries, or past its 15 minutes. Nothing about the address is kept after that.
CREATE TABLE email_codes (
  email TEXT PRIMARY KEY,
  code_hash TEXT NOT NULL,
  expires_at INTEGER NOT NULL,
  tries INTEGER NOT NULL DEFAULT 0
);
