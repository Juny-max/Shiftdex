-- Apply once to the existing production D1 database before deploying the hardened Worker.
ALTER TABLE worker_accounts ADD COLUMN password_iterations INTEGER NOT NULL DEFAULT 100000;
CREATE TABLE IF NOT EXISTS worker_auth_limits (
  scope TEXT NOT NULL,
  email TEXT NOT NULL,
  ip_hash TEXT NOT NULL,
  failures INTEGER NOT NULL DEFAULT 0,
  window_started_at TEXT NOT NULL,
  blocked_until TEXT,
  PRIMARY KEY(scope,email,ip_hash)
);

-- Clear old bearer sessions after moving browser tokens to per-tab session storage.
DELETE FROM worker_sessions;
DELETE FROM admin_sessions;
