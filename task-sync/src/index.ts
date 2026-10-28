// UnifyTasks: Google Tasks <-> Microsoft To Do two way sync
// Part 1: OAuth login, token storage and refresh
// Part 2: list pairing, sync engine, cron trigger, logs

type AppEnv = Env & {
  GOOGLE_CLIENT_ID: string;
  GOOGLE_CLIENT_SECRET: string;
  MS_CLIENT_ID: string;
  MS_CLIENT_SECRET: string;
  ADMIN_KEY: string;
};

type Provider = "google" | "microsoft";

type TokenResponse = {
  access_token: string;
  refresh_token?: string;
  expires_in: number;
};

// Max create/update/delete calls per run. Keeps us under the free plan's
// 50 outgoing requests per invocation. Anything left over runs next minute.
const WRITE_BUDGET = 35;

// How long a task must stay missing before its twin is deleted (two step delete)
const DELETE_DELAY_MS = 30_000;

const PROVIDERS: Record<
  Provider,
  { authUrl: string; tokenUrl: string; scope: string; extraAuthParams: Record<string, string> }
> = {
  google: {
    authUrl: "https://accounts.google.com/o/oauth2/v2/auth",
    tokenUrl: "https://oauth2.googleapis.com/token",
    scope: "https://www.googleapis.com/auth/tasks",
    extraAuthParams: { access_type: "offline", prompt: "consent" },
  },
  microsoft: {
    authUrl: "https://login.microsoftonline.com/common/oauth2/v2.0/authorize",
    tokenUrl: "https://login.microsoftonline.com/common/oauth2/v2.0/token",
    scope: "Tasks.ReadWrite offline_access User.Read",
    extraAuthParams: { prompt: "select_account" },
  },
};

// ======================================================================
// Helpers
// ======================================================================

function creds(env: AppEnv, p: Provider) {
  return p === "google"
    ? { id: env.GOOGLE_CLIENT_ID, secret: env.GOOGLE_CLIENT_SECRET }
    : { id: env.MS_CLIENT_ID, secret: env.MS_CLIENT_SECRET };
}

// Built from the incoming request, so the same code works on any Worker URL
function redirectUri(req: Request, p: Provider): string {
  return `${new URL(req.url).origin}/auth/${p}/callback`;
}

function isAdmin(req: Request, env: AppEnv): boolean {
  return new URL(req.url).searchParams.get("key") === env.ADMIN_KEY;
}

function getCookie(req: Request, name: string): string | null {
  const header = req.headers.get("Cookie") || "";
  for (const part of header.split(";")) {
    const [k, ...v] = part.trim().split("=");
    if (k === name) return v.join("=");
  }
  return null;
}

function json(data: unknown, status = 200): Response {
  return new Response(JSON.stringify(data, null, 2), {
    status,
    headers: { "Content-Type": "application/json" },
  });
}

function logStmt(env: AppEnv, level: string, message: string): D1PreparedStatement {
  return env.DB.prepare("INSERT INTO sync_log (ts, level, message) VALUES (?, ?, ?)").bind(
    Date.now(),
    level,
    message.slice(0, 2000)
  );
}

async function log(env: AppEnv, level: string, message: string): Promise<void> {
  await logStmt(env, level, message).run();
}

// ======================================================================
// Token storage and refresh
// ======================================================================

async function saveNewTokens(env: AppEnv, p: Provider, t: TokenResponse): Promise<void> {
  const expiresAt = Date.now() + (t.expires_in - 60) * 1000;
  await env.DB.prepare(
    "INSERT OR REPLACE INTO tokens (provider, access_token, refresh_token, expires_at) VALUES (?, ?, ?, ?)"
  )
    .bind(p, t.access_token, t.refresh_token, expiresAt)
    .run();
}

async function updateRefreshedTokens(env: AppEnv, p: Provider, t: TokenResponse): Promise<void> {
  const expiresAt = Date.now() + (t.expires_in - 60) * 1000;
  if (t.refresh_token) {
    // Microsoft rotates refresh tokens, so keep the newest one
    await env.DB.prepare(
      "UPDATE tokens SET access_token = ?, refresh_token = ?, expires_at = ? WHERE provider = ?"
    )
      .bind(t.access_token, t.refresh_token, expiresAt, p)
      .run();
  } else {
    await env.DB.prepare("UPDATE tokens SET access_token = ?, expires_at = ? WHERE provider = ?")
      .bind(t.access_token, expiresAt, p)
      .run();
  }
}

async function getAccessToken(env: AppEnv, p: Provider): Promise<string> {
  const row = await env.DB.prepare(
    "SELECT access_token, refresh_token, expires_at FROM tokens WHERE provider = ?"
  )
    .bind(p)
    .first<{ access_token: string; refresh_token: string; expires_at: number }>();

  if (!row) throw new Error(`${p} is not connected yet`);
  if (row.access_token && row.expires_at > Date.now()) return row.access_token;

  const { id, secret } = creds(env, p);
  const body = new URLSearchParams({
    client_id: id,
    client_secret: secret,
    refresh_token: row.refresh_token,
    grant_type: "refresh_token",
  });
  if (p === "microsoft") body.set("scope", PROVIDERS.microsoft.scope);

  const res = await fetch(PROVIDERS[p].tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });
  const data = (await res.json()) as TokenResponse & { error?: string };
  if (!res.ok || !data.access_token) {
    await log(env, "error", `${p} token refresh failed: ${JSON.stringify(data)}`);
    throw new Error(`${p} token refresh failed, sign in again`);
  }
  await updateRefreshedTokens(env, p, data);
  return data.access_token;
}

// ======================================================================
// OAuth routes
// ======================================================================

function startAuth(req: Request, env: AppEnv, p: Provider): Response {
  if (!isAdmin(req, env)) return new Response("Forbidden", { status: 403 });

  const state = crypto.randomUUID();
  const cfg = PROVIDERS[p];
  const params = new URLSearchParams({
    client_id: creds(env, p).id,
    redirect_uri: redirectUri(req, p),
    response_type: "code",
    scope: cfg.scope,
    state,
    ...cfg.extraAuthParams,
  });

  return new Response(null, {
    status: 302,
    headers: {
      Location: `${cfg.authUrl}?${params.toString()}`,
      "Set-Cookie": `oauth_state=${state}; Path=/auth; HttpOnly; Secure; SameSite=Lax; Max-Age=600`,
    },
  });
}

async function handleCallback(req: Request, env: AppEnv, p: Provider): Promise<Response> {
  const url = new URL(req.url);
  const error = url.searchParams.get("error");
  if (error) return new Response(`Sign in failed: ${error}`, { status: 400 });

  const code = url.searchParams.get("code");
  const state = url.searchParams.get("state");
  if (!code || !state || state !== getCookie(req, "oauth_state")) {
    return new Response("Invalid or expired login attempt. Start again from the login link.", {
      status: 400,
    });
  }

  const { id, secret } = creds(env, p);
  const res = await fetch(PROVIDERS[p].tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: id,
      client_secret: secret,
      code,
      redirect_uri: redirectUri(req, p),
      grant_type: "authorization_code",
    }),
  });
  const data = (await res.json()) as TokenResponse & { error?: string };

  if (!res.ok || !data.refresh_token) {
    await log(env, "error", `${p} token exchange failed: ${JSON.stringify(data)}`);
    return new Response(`Token exchange failed for ${p}. Check /logs.`, { status: 500 });
  }

  await saveNewTokens(env, p, data);
  await log(env, "info", `${p} connected`);
  return new Response(`${p} connected. You can close this tab.`, {
    headers: { "Set-Cookie": "oauth_state=; Path=/auth; Max-Age=0" },
  });
}

// ======================================================================
// Task shapes and normalization
// ======================================================================

type GTask = {
  id: string;
  title?: string;
  notes?: string;
  due?: string;
  status?: "needsAction" | "completed";
  updated?: string;
  deleted?: boolean;
  hidden?: boolean;
};

type MTask = {
  id: string;
  title?: string;
  body?: { content?: string; contentType?: string };
  dueDateTime?: { dateTime: string; timeZone: string } | null;
  status?: string;
  lastModifiedDateTime?: string;
};

// The shared fields both services support. Everything is compared in this shape.
type Canon = { title: string; notes: string; due: string | null; completed: boolean };

function stripHtml(s: string): string {
  return s
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<[^>]*>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .trim();
}

function canonG(g: GTask): Canon {
  return {
    title: (g.title || "").trim(),
    notes: (g.notes || "").trim(),
    // Google stores due as a date only, so keep just YYYY-MM-DD
    due: g.due ? g.due.slice(0, 10) : null,
    completed: g.status === "completed",
  };
}

function canonM(m: MTask): Canon {
  const raw = m.body?.content || "";
  const notes = m.body?.contentType === "html" ? stripHtml(raw) : raw.trim();
  return {
    title: (m.title || "").trim(),
    notes,
    due: m.dueDateTime?.dateTime ? m.dueDateTime.dateTime.slice(0, 10) : null,
    completed: m.status === "completed",
  };
}

async function hashCanon(c: Canon): Promise<string> {
  const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(JSON.stringify(c)));
  return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, "0")).join("");
}

function toGoogleBody(c: Canon): Record<string, unknown> {
  const body: Record<string, unknown> = {
    title: c.title,
    notes: c.notes,
    due: c.due ? `${c.due}T00:00:00.000Z` : null,
    status: c.completed ? "completed" : "needsAction",
  };
  // Google needs the completed timestamp cleared when a task is reopened
  if (!c.completed) body.completed = null;
  return body;
}

function toMsBody(c: Canon): Record<string, unknown> {
  return {
    title: c.title,
    body: { content: c.notes, contentType: "text" },
    dueDateTime: c.due ? { dateTime: `${c.due}T00:00:00`, timeZone: "UTC" } : null,
    status: c.completed ? "completed" : "notStarted",
  };
}

// ======================================================================
// API calls
// ======================================================================

type Stats = {
  createdInGoogle: number;
  createdInMicrosoft: number;
  updatedGoogle: number;
  updatedMicrosoft: number;
  deletedGoogle: number;
  deletedMicrosoft: number;
  linkedExisting: number;
  deferred: number;
};

type Ctx = {
  env: AppEnv;
  gToken: string;
  mToken: string;
  writes: number;
  stats: Stats;
};

function canWrite(ctx: Ctx): boolean {
  if (ctx.writes < WRITE_BUDGET) return true;
  ctx.stats.deferred++;
  return false;
}

async function api(
  token: string,
  label: string,
  url: string,
  init: RequestInit = {}
): Promise<any> {
  const res = await fetch(url, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
  });
  const method = init.method || "GET";
  // Deleting something that is already gone counts as success
  if (method === "DELETE" && res.status === 404) return null;
  if (!res.ok) throw new Error(`${label} ${method} ${res.status}: ${(await res.text()).slice(0, 500)}`);
  if (res.status === 204) return null;
  return res.json();
}

const G = "https://tasks.googleapis.com/tasks/v1";
const M = "https://graph.microsoft.com/v1.0/me/todo";

async function listGoogle(ctx: Ctx, listId: string): Promise<GTask[]> {
  const out: GTask[] = [];
  let pageToken = "";
  do {
    const params = new URLSearchParams({
      maxResults: "100",
      showCompleted: "true",
      showHidden: "true",
      showDeleted: "true",
    });
    if (pageToken) params.set("pageToken", pageToken);
    const data = await api(ctx.gToken, "Google", `${G}/lists/${listId}/tasks?${params}`);
    out.push(...(data.items || []));
    pageToken = data.nextPageToken || "";
  } while (pageToken);
  return out;
}

async function listMicrosoft(ctx: Ctx, listId: string): Promise<MTask[]> {
  const out: MTask[] = [];
  let url = `${M}/lists/${listId}/tasks?$top=100`;
  while (url) {
    const data = await api(ctx.mToken, "Microsoft", url);
    out.push(...(data.value || []));
    url = data["@odata.nextLink"] || "";
  }
  return out;
}

async function googleCreate(ctx: Ctx, listId: string, c: Canon): Promise<GTask> {
  ctx.writes++;
  ctx.stats.createdInGoogle++;
  return api(ctx.gToken, "Google", `${G}/lists/${listId}/tasks`, {
    method: "POST",
    body: JSON.stringify(toGoogleBody(c)),
  });
}

async function googleUpdate(ctx: Ctx, listId: string, id: string, c: Canon): Promise<void> {
  ctx.writes++;
  ctx.stats.updatedGoogle++;
  await api(ctx.gToken, "Google", `${G}/lists/${listId}/tasks/${id}`, {
    method: "PATCH",
    body: JSON.stringify(toGoogleBody(c)),
  });
}

async function googleDelete(ctx: Ctx, listId: string, id: string): Promise<void> {
  ctx.writes++;
  ctx.stats.deletedGoogle++;
  await api(ctx.gToken, "Google", `${G}/lists/${listId}/tasks/${id}`, { method: "DELETE" });
}

async function msCreate(ctx: Ctx, listId: string, c: Canon): Promise<MTask> {
  ctx.writes++;
  ctx.stats.createdInMicrosoft++;
  return api(ctx.mToken, "Microsoft", `${M}/lists/${listId}/tasks`, {
    method: "POST",
    body: JSON.stringify(toMsBody(c)),
  });
}

async function msUpdate(ctx: Ctx, listId: string, id: string, c: Canon): Promise<void> {
  ctx.writes++;
  ctx.stats.updatedMicrosoft++;
  await api(ctx.mToken, "Microsoft", `${M}/lists/${listId}/tasks/${id}`, {
    method: "PATCH",
    body: JSON.stringify(toMsBody(c)),
  });
}

async function msDelete(ctx: Ctx, listId: string, id: string): Promise<void> {
  ctx.writes++;
  ctx.stats.deletedMicrosoft++;
  await api(ctx.mToken, "Microsoft", `${M}/lists/${listId}/tasks/${id}`, { method: "DELETE" });
}

// ======================================================================
// Sync engine
// ======================================================================

type Pair = { id: number; google_list_id: string; ms_list_id: string };

type MapRow = {
  id: number;
  google_task_id: string;
  ms_task_id: string;
  last_hash: string | null;
  deleted_at: number | null;
};

async function syncPair(ctx: Ctx, pair: Pair, stmts: D1PreparedStatement[]): Promise<void> {
  const { env } = ctx;
  const gList = pair.google_list_id;
  const mList = pair.ms_list_id;

  const [gTasks, mTasks, mapResult] = await Promise.all([
    listGoogle(ctx, gList),
    listMicrosoft(ctx, mList),
    env.DB.prepare(
      "SELECT id, google_task_id, ms_task_id, last_hash, deleted_at FROM task_map WHERE pair_id = ?"
    )
      .bind(pair.id)
      .all<MapRow>(),
  ]);
  const maps = mapResult.results;

  const gById = new Map(gTasks.filter((t) => !t.deleted).map((t) => [t.id, t]));
  const mById = new Map(mTasks.map((t) => [t.id, t]));
  const mappedG = new Set<string>();
  const mappedM = new Set<string>();
  const now = Date.now();

  // 1. Existing pairs: propagate edits and deletions
  for (const row of maps) {
    mappedG.add(row.google_task_id);
    mappedM.add(row.ms_task_id);
    const g = gById.get(row.google_task_id);
    const m = mById.get(row.ms_task_id);

    if (!g && !m) {
      stmts.push(env.DB.prepare("DELETE FROM task_map WHERE id = ?").bind(row.id));
      continue;
    }

    if (!g || !m) {
      // Two step delete: first notice, then delete the twin on a later run
      if (!row.deleted_at) {
        stmts.push(env.DB.prepare("UPDATE task_map SET deleted_at = ? WHERE id = ?").bind(now, row.id));
        continue;
      }
      if (now - row.deleted_at < DELETE_DELAY_MS) continue;
      if (!canWrite(ctx)) continue;
      if (g) await googleDelete(ctx, gList, g.id);
      else await msDelete(ctx, mList, m!.id);
      stmts.push(env.DB.prepare("DELETE FROM task_map WHERE id = ?").bind(row.id));
      continue;
    }

    // Both exist. If one had been marked missing and came back, clear the mark.
    if (row.deleted_at) {
      stmts.push(env.DB.prepare("UPDATE task_map SET deleted_at = NULL WHERE id = ?").bind(row.id));
    }

    const cg = canonG(g);
    const cm = canonM(m);
    const hg = await hashCanon(cg);
    const hm = await hashCanon(cm);

    if (hg === hm) {
      if (row.last_hash !== hg) {
        stmts.push(env.DB.prepare("UPDATE task_map SET last_hash = ? WHERE id = ?").bind(hg, row.id));
      }
      continue;
    }

    if (!canWrite(ctx)) continue;

    // Decide which side changed since the last sync
    let googleWins: boolean;
    if (hm === row.last_hash) googleWins = true;
    else if (hg === row.last_hash) googleWins = false;
    else {
      // Both changed: the most recently edited one wins
      googleWins = Date.parse(g.updated || "0") >= Date.parse(m.lastModifiedDateTime || "0");
    }

    if (googleWins) {
      await msUpdate(ctx, mList, m.id, cg);
      stmts.push(env.DB.prepare("UPDATE task_map SET last_hash = ? WHERE id = ?").bind(hg, row.id));
    } else {
      await googleUpdate(ctx, gList, g.id, cm);
      stmts.push(env.DB.prepare("UPDATE task_map SET last_hash = ? WHERE id = ?").bind(hm, row.id));
    }
  }

  // 2. New tasks on either side
  const newG = gTasks.filter((t) => !t.deleted && !mappedG.has(t.id) && (t.title || "").trim());
  const newM = mTasks.filter((t) => !mappedM.has(t.id) && (t.title || "").trim());

  // Link tasks with the same title instead of duplicating them (covers the first import)
  const mByTitle = new Map<string, MTask[]>();
  for (const m of newM) {
    const key = canonM(m).title.toLowerCase();
    if (!mByTitle.has(key)) mByTitle.set(key, []);
    mByTitle.get(key)!.push(m);
  }
  const linkedM = new Set<string>();

  const insertMap = (gid: string, mid: string, hash: string | null) =>
    env.DB.prepare(
      "INSERT INTO task_map (pair_id, google_task_id, ms_task_id, last_hash) VALUES (?, ?, ?, ?)"
    ).bind(pair.id, gid, mid, hash);

  for (const g of newG) {
    const cg = canonG(g);
    const match = mByTitle.get(cg.title.toLowerCase())?.find((m) => !linkedM.has(m.id));
    if (match) {
      linkedM.add(match.id);
      ctx.stats.linkedExisting++;
      const hg = await hashCanon(cg);
      const hm = await hashCanon(canonM(match));
      // If contents differ, a null hash makes the next run pick the newest version
      stmts.push(insertMap(g.id, match.id, hg === hm ? hg : null));
      continue;
    }
    if (!canWrite(ctx)) continue;
    const created = await msCreate(ctx, mList, cg);
    stmts.push(insertMap(g.id, created.id, await hashCanon(cg)));
  }

  for (const m of newM) {
    if (linkedM.has(m.id)) continue;
    if (!canWrite(ctx)) continue;
    const cm = canonM(m);
    const created = await googleCreate(ctx, gList, cm);
    stmts.push(insertMap(created.id, m.id, await hashCanon(cm)));
  }

  stmts.push(
    env.DB.prepare("UPDATE list_pairs SET google_last_poll = ? WHERE id = ?").bind(
      new Date().toISOString(),
      pair.id
    )
  );
}

async function runSync(env: AppEnv): Promise<Stats> {
  const ctx: Ctx = {
    env,
    gToken: await getAccessToken(env, "google"),
    mToken: await getAccessToken(env, "microsoft"),
    writes: 0,
    stats: {
      createdInGoogle: 0,
      createdInMicrosoft: 0,
      updatedGoogle: 0,
      updatedMicrosoft: 0,
      deletedGoogle: 0,
      deletedMicrosoft: 0,
      linkedExisting: 0,
      deferred: 0,
    },
  };

  const { results: pairs } = await env.DB.prepare(
    "SELECT id, google_list_id, ms_list_id FROM list_pairs"
  ).all<Pair>();

  const stmts: D1PreparedStatement[] = [];
  try {
    for (const pair of pairs) await syncPair(ctx, pair, stmts);
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    stmts.push(logStmt(env, "error", `sync failed: ${message}`));
    throw err;
  } finally {
    // Always save what was done, even if a later step failed,
    // so tasks created this run are not created again next run
    if (ctx.writes > 0) stmts.push(logStmt(env, "info", `sync: ${JSON.stringify(ctx.stats)}`));
    if (stmts.length) await env.DB.batch(stmts);
  }
  return ctx.stats;
}

// ======================================================================
// Setup, status, logs and test routes
// ======================================================================

async function setupDefaultPair(env: AppEnv): Promise<Response> {
  const gToken = await getAccessToken(env, "google");
  const mToken = await getAccessToken(env, "microsoft");

  const gData = await api(gToken, "Google", `${G}/users/@me/lists`);
  const mData = await api(mToken, "Microsoft", `${M}/lists`);

  const gDefault = (gData.items || [])[0] as { id: string; title: string } | undefined;
  const mDefault = (mData.value || []).find(
    (l: { wellknownListName?: string }) => l.wellknownListName === "defaultList"
  ) as { id: string; displayName: string } | undefined;

  if (!gDefault || !mDefault) return json({ error: "Could not find default lists" }, 500);

  await env.DB.prepare(
    "INSERT OR IGNORE INTO list_pairs (google_list_id, ms_list_id) VALUES (?, ?)"
  )
    .bind(gDefault.id, mDefault.id)
    .run();
  await log(env, "info", `paired "${gDefault.title}" with "${mDefault.displayName}"`);

  return json({ paired: { google: gDefault.title, microsoft: mDefault.displayName } });
}

async function status(env: AppEnv): Promise<Response> {
  const { results: tokens } = await env.DB.prepare("SELECT provider FROM tokens").all<{
    provider: string;
  }>();
  const pairs = await env.DB.prepare("SELECT COUNT(*) AS n FROM list_pairs").first<{ n: number }>();
  const tasks = await env.DB.prepare("SELECT COUNT(*) AS n FROM task_map").first<{ n: number }>();
  const lastPoll = await env.DB.prepare(
    "SELECT MAX(google_last_poll) AS t FROM list_pairs"
  ).first<{ t: string | null }>();
  return json({
    google: tokens.some((r) => r.provider === "google") ? "connected" : "not connected",
    microsoft: tokens.some((r) => r.provider === "microsoft") ? "connected" : "not connected",
    list_pairs: pairs?.n ?? 0,
    synced_tasks: tasks?.n ?? 0,
    last_sync: lastPoll?.t ?? "never",
  });
}

async function recentLogs(env: AppEnv): Promise<Response> {
  const { results } = await env.DB.prepare(
    "SELECT ts, level, message FROM sync_log ORDER BY id DESC LIMIT 30"
  ).all<{ ts: number; level: string; message: string }>();
  return json(results.map((r) => ({ time: new Date(r.ts).toISOString(), level: r.level, message: r.message })));
}

async function testGoogle(env: AppEnv): Promise<Response> {
  const token = await getAccessToken(env, "google");
  const data = await api(token, "Google", `${G}/users/@me/lists`);
  return json({
    google_task_lists: (data.items || []).map((l: { id: string; title: string }) => ({
      id: l.id,
      title: l.title,
    })),
  });
}

async function testMicrosoft(env: AppEnv): Promise<Response> {
  const token = await getAccessToken(env, "microsoft");
  const data = await api(token, "Microsoft", `${M}/lists`);
  return json({
    microsoft_todo_lists: (data.value || []).map((l: { id: string; displayName: string }) => ({
      id: l.id,
      title: l.displayName,
    })),
  });
}

// ======================================================================
// Entry points
// ======================================================================

export default {
  async fetch(req: Request, env: AppEnv): Promise<Response> {
    const path = new URL(req.url).pathname;

    try {
      if (path === "/auth/google") return startAuth(req, env, "google");
      if (path === "/auth/microsoft") return startAuth(req, env, "microsoft");
      if (path === "/auth/google/callback") return await handleCallback(req, env, "google");
      if (path === "/auth/microsoft/callback") return await handleCallback(req, env, "microsoft");

      if (path === "/status") return await status(env);

      const adminRoutes: Record<string, () => Promise<Response>> = {
        "/setup": () => setupDefaultPair(env),
        "/sync": async () => json(await runSync(env)),
        "/logs": () => recentLogs(env),
        "/test/google": () => testGoogle(env),
        "/test/microsoft": () => testMicrosoft(env),
      };
      if (adminRoutes[path]) {
        if (!isAdmin(req, env)) return new Response("Forbidden", { status: 403 });
        return await adminRoutes[path]();
      }

      return new Response("UnifyTasks is running.");
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      await log(env, "error", `${path}: ${message}`);
      return new Response(`Error: ${message}`, { status: 500 });
    }
  },

  // Runs every minute (configured in wrangler.jsonc)
  async scheduled(_controller: ScheduledController, env: AppEnv, ctx: ExecutionContext): Promise<void> {
    ctx.waitUntil(
      runSync(env).catch(async (err) => {
        const message = err instanceof Error ? err.message : String(err);
        await log(env, "error", `cron: ${message}`);
      })
    );
  },
} satisfies ExportedHandler<AppEnv>;
