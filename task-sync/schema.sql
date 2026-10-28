CREATE TABLE IF NOT EXISTS tokens (
  provider TEXT PRIMARY KEY,
  access_token TEXT,
  refresh_token TEXT NOT NULL,
  expires_at INTEGER NOT NULL
);

CREATE TABLE IF NOT EXISTS list_pairs (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  google_list_id TEXT NOT NULL UNIQUE,
  ms_list_id TEXT NOT NULL UNIQUE,
  google_last_poll TEXT,
  ms_delta_link TEXT,
  ms_subscription_id TEXT,
  ms_subscription_expires INTEGER
);

CREATE TABLE IF NOT EXISTS task_map (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  pair_id INTEGER NOT NULL REFERENCES list_pairs(id),
  google_task_id TEXT UNIQUE,
  ms_task_id TEXT UNIQUE,
  last_hash TEXT,
  google_updated TEXT,
  ms_updated TEXT,
  deleted_at INTEGER
);

CREATE TABLE IF NOT EXISTS sync_log (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  ts INTEGER NOT NULL,
  level TEXT NOT NULL,
  message TEXT NOT NULL
);