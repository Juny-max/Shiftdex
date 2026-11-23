-- 0003_worker_passkeys.sql
-- Production D1 tables for WebAuthn/FIDO2 Passkeys & Challenge validation

CREATE TABLE IF NOT EXISTS worker_passkeys (
  id TEXT PRIMARY KEY,
  account_id TEXT NOT NULL REFERENCES worker_accounts(id) ON DELETE CASCADE,
  credential_id TEXT NOT NULL UNIQUE,
  public_key TEXT NOT NULL,
  counter INTEGER NOT NULL DEFAULT 0,
  device_name TEXT,
  transports TEXT,
  backup_eligible INTEGER NOT NULL DEFAULT 0,
  backup_status INTEGER NOT NULL DEFAULT 0,
  created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
  last_used_at TEXT
);

CREATE INDEX IF NOT EXISTS worker_passkeys_account_idx ON worker_passkeys(account_id);
CREATE INDEX IF NOT EXISTS worker_passkeys_credential_idx ON worker_passkeys(credential_id);

CREATE TABLE IF NOT EXISTS worker_auth_challenges (
  challenge_id TEXT PRIMARY KEY,
  challenge TEXT NOT NULL,
  account_id TEXT,
  flow_type TEXT NOT NULL CHECK(flow_type IN ('registration', 'authentication')),
  expires_at TEXT NOT NULL,
  created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS auth_challenges_expiry_idx ON worker_auth_challenges(expires_at);
