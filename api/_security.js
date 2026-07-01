import { createHash } from 'node:crypto';
import { createClient } from '@supabase/supabase-js';

export const ACCESS_COOKIE = 'cf_access';
export const REFRESH_COOKIE = 'cf_refresh';

function env(name, fallbackName) {
  const value = process.env[name] || (fallbackName ? process.env[fallbackName] : undefined);
  if (!value) throw new Error(`Missing server environment variable: ${name}`);
  return value;
}

export function supabaseConfig() {
  return {
    url: env('SUPABASE_URL', 'VITE_SUPABASE_URL'),
    anonKey: env('SUPABASE_ANON_KEY', 'VITE_SUPABASE_ANON_KEY'),
  };
}

export function createSupabase(accessToken) {
  const { url, anonKey } = supabaseConfig();
  return createClient(url, anonKey, {
    auth: { persistSession: false, autoRefreshToken: false, detectSessionInUrl: false },
    global: accessToken ? { headers: { Authorization: `Bearer ${accessToken}` } } : undefined,
  });
}

function parseCookies(req) {
  const header = req.headers.cookie || '';
  return Object.fromEntries(header.split(';').flatMap((part) => {
    const separator = part.indexOf('=');
    if (separator < 0) return [];
    const key = part.slice(0, separator).trim();
    const value = part.slice(separator + 1).trim();
    try { return [[key, decodeURIComponent(value)]]; } catch { return []; }
  }));
}

function appendCookie(res, cookie) {
  const current = res.getHeader('Set-Cookie');
  const next = current ? (Array.isArray(current) ? [...current, cookie] : [current, cookie]) : [cookie];
  res.setHeader('Set-Cookie', next);
}

function serializeCookie(name, value, maxAge) {
  const secure = process.env.VERCEL_ENV === 'production' || process.env.NODE_ENV === 'production' ? '; Secure' : '';
  return `${name}=${encodeURIComponent(value)}; Path=/; Max-Age=${maxAge}; HttpOnly${secure}; SameSite=Strict`;
}

export function setAuthCookies(res, session) {
  const accessMaxAge = Math.max(60, Number(session.expires_in || 3600));
  appendCookie(res, serializeCookie(ACCESS_COOKIE, session.access_token, accessMaxAge));
  appendCookie(res, serializeCookie(REFRESH_COOKIE, session.refresh_token, 60 * 60 * 24 * 30));
  res.setHeader('Cache-Control', 'no-store');
}

export function clearAuthCookies(res) {
  appendCookie(res, serializeCookie(ACCESS_COOKIE, '', 0));
  appendCookie(res, serializeCookie(REFRESH_COOKIE, '', 0));
  res.setHeader('Cache-Control', 'no-store');
}

export function publicSession(session) {
  if (!session?.user) return null;
  return {
    user: {
      id: session.user.id,
      email: session.user.email || null,
      user_metadata: session.user.user_metadata || {},
    },
  };
}

export async function getAuthenticatedContext(req, res) {
  const cookies = parseCookies(req);
  let accessToken = cookies[ACCESS_COOKIE];
  const refreshToken = cookies[REFRESH_COOKIE];
  let client = createSupabase(accessToken);

  if (accessToken) {
    const { data } = await client.auth.getUser(accessToken);
    if (data.user) return { client, user: data.user, accessToken, refreshToken };
  }

  if (!refreshToken) return null;
  const refreshClient = createSupabase();
  const { data, error } = await refreshClient.auth.refreshSession({ refresh_token: refreshToken });
  if (error || !data.session) {
    clearAuthCookies(res);
    return null;
  }

  setAuthCookies(res, data.session);
  accessToken = data.session.access_token;
  client = createSupabase(accessToken);
  return { client, user: data.session.user, accessToken, refreshToken: data.session.refresh_token };
}

export async function revokeSession(context) {
  const serviceRoleKey = process.env.SUPABASE_SERVICE_ROLE_KEY;
  if (serviceRoleKey) {
    const { url } = supabaseConfig();
    const admin = createClient(url, serviceRoleKey, {
      auth: { persistSession: false, autoRefreshToken: false, detectSessionInUrl: false },
    });
    await admin.auth.admin.signOut(context.accessToken, 'local');
    return;
  }

  if (context.refreshToken) {
    const client = createSupabase();
    const { error } = await client.auth.setSession({
      access_token: context.accessToken,
      refresh_token: context.refreshToken,
    });
    if (!error) await client.auth.signOut({ scope: 'local' });
  }
}

export function assertSameOrigin(req) {
  const origin = req.headers.origin;
  const host = req.headers['x-forwarded-host'] || req.headers.host;
  if (!origin || !host) return false;
  try { return new URL(origin).host === host; } catch { return false; }
}

export function sendJson(res, status, body) {
  res.statusCode = status;
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.setHeader('Cache-Control', 'no-store');
  res.end(JSON.stringify(body));
}

export function requestFingerprint(req, scope, discriminator = '') {
  const forwarded = String(req.headers['x-forwarded-for'] || '').split(',')[0].trim();
  const ip = forwarded || req.socket?.remoteAddress || 'unknown';
  return createHash('sha256').update(`${scope}:${ip}:${discriminator.toLowerCase()}`).digest('hex');
}

const localLimits = new Map();

export async function enforceRateLimit(key, limit, windowSeconds) {
  const serviceRoleKey = process.env.SUPABASE_SERVICE_ROLE_KEY;
  if (serviceRoleKey) {
    try {
      const { url } = supabaseConfig();
      const admin = createClient(url, serviceRoleKey, {
        auth: { persistSession: false, autoRefreshToken: false, detectSessionInUrl: false },
      });
      const { data, error } = await admin.rpc('consume_rate_limit', {
        p_key: key,
        p_limit: limit,
        p_window_seconds: windowSeconds,
      });
      if (!error) return data === true;
      console.error('[rate-limit] RPC error, falling back to memory:', error.message);
    } catch (err) {
      console.error('[rate-limit] unexpected error, falling back to memory:', err);
    }
  }

  if (process.env.VERCEL_ENV === 'production' && !serviceRoleKey) {
    throw new Error('SUPABASE_SERVICE_ROLE_KEY is required in production');
  }

  const now = Date.now();
  const entry = localLimits.get(key);
  if (!entry || entry.resetAt <= now) {
    localLimits.set(key, { count: 1, resetAt: now + windowSeconds * 1000 });
    return true;
  }
  entry.count += 1;
  return entry.count <= limit;
}

export function createServiceRoleClient() {
  const { url } = supabaseConfig();
  const key = process.env.SUPABASE_SERVICE_ROLE_KEY;
  if (!key) throw new Error('Missing SUPABASE_SERVICE_ROLE_KEY');
  return createClient(url, key, {
    auth: { persistSession: false, autoRefreshToken: false, detectSessionInUrl: false },
  });
}

export function requestOrigin(req) {
  const protocol = req.headers['x-forwarded-proto'] || 'https';
  const host = req.headers['x-forwarded-host'] || req.headers.host;
  if (!host) throw new Error('Missing host');
  return `${protocol}://${host}`;
}
