import { z } from 'zod';
import {
  assertSameOrigin,
  clearAuthCookies,
  createSupabase,
  enforceRateLimit,
  ensureUserProfile,
  getAuthenticatedContext,
  publicSession,
  requestFingerprint,
  requestOrigin,
  revokeSession,
  sendJson,
  setAuthCookies,
} from './_security.js';

const email = z.string().trim().email().max(254);
const fullName = z.string().trim().min(2).max(160);
const loginPassword = z.string().min(1).max(128);
const newPassword = z.string().min(8).max(128);
const requestSchema = z.discriminatedUnion('action', [
  z.object({ action: z.literal('login'), email, password: loginPassword }),
  z.object({ action: z.literal('signup'), email, password: newPassword, fullName }),
  z.object({ action: z.literal('recovery'), email }),
  z.object({ action: z.literal('establish-recovery'), accessToken: z.string().min(20).max(4096), refreshToken: z.string().min(20).max(4096) }),
  z.object({ action: z.literal('update-password'), password: newPassword }),
  z.object({ action: z.literal('logout') }),
]);

function safeAuthError(error) {
  const message = String(error?.message || '').toLowerCase();
  if (message.includes('invalid login credentials')) return 'Invalid login credentials';
  if (message.includes('email not confirmed')) return 'Email not confirmed';
  if (message.includes('already registered')) return 'User already registered';
  if (message.includes('password')) return 'Password does not meet security requirements';
  if (message.includes('email')) return 'Invalid email';
  return 'Authentication request failed';
}

export default async function handler(req, res) {
  try {
    if (req.method === 'GET') {
      const context = await getAuthenticatedContext(req, res);
      return sendJson(res, 200, { session: context ? publicSession({ user: context.user }) : null });
    }

    if (req.method !== 'POST') return sendJson(res, 405, { error: 'Method not allowed' });
    if (!assertSameOrigin(req)) return sendJson(res, 403, { error: 'Invalid request origin' });

    const parsed = requestSchema.safeParse(req.body);
    if (!parsed.success) return sendJson(res, 400, { error: 'Invalid request data' });
    const body = parsed.data;
    const identity = 'email' in body ? body.email : '';
    const limits = {
      login: [5, 15 * 60],
      signup: [3, 60 * 60],
      recovery: [3, 60 * 60],
      'establish-recovery': [5, 15 * 60],
      'update-password': [5, 60 * 60],
      logout: [30, 60],
    };
    const [limit, windowSeconds] = limits[body.action];
    const allowed = await enforceRateLimit(requestFingerprint(req, body.action, identity), limit, windowSeconds);
    if (!allowed) return sendJson(res, 429, { error: 'Too many requests. Try again later.' });

    if (body.action === 'logout') {
      const context = await getAuthenticatedContext(req, res);
      clearAuthCookies(res);
      if (context) {
        try { await revokeSession(context); }
        catch (error) { console.error('Session revocation failed', error); }
      }
      return sendJson(res, 200, { ok: true });
    }

    if (body.action === 'update-password') {
      const context = await getAuthenticatedContext(req, res);
      if (!context) return sendJson(res, 401, { error: 'Unauthorized' });
      const { error } = await context.client.auth.updateUser({ password: body.password });
      if (error) return sendJson(res, 400, { error: safeAuthError(error) });
      return sendJson(res, 200, { ok: true });
    }

    const supabase = createSupabase();

    if (body.action === 'login') {
      const { data, error } = await supabase.auth.signInWithPassword({ email: body.email, password: body.password });
      if (error || !data.session) return sendJson(res, 401, { error: safeAuthError(error) });
      try { await ensureUserProfile(data.session); }
      catch (profileError) { console.error('Profile repair after login failed', profileError); }
      setAuthCookies(res, data.session);
      return sendJson(res, 200, { session: publicSession(data.session) });
    }

    if (body.action === 'signup') {
      const { data, error } = await supabase.auth.signUp({
        email: body.email,
        password: body.password,
        options: {
          emailRedirectTo: requestOrigin(req),
          data: { full_name: body.fullName },
        },
      });
      if (error) return sendJson(res, 400, { error: safeAuthError(error) });
      const profileContext = data.session || (data.user ? { user: data.user } : null);
      if (profileContext) {
        try { await ensureUserProfile(profileContext); }
        catch (profileError) { console.error('Profile repair after signup failed', profileError); }
      }
      if (data.session) setAuthCookies(res, data.session);
      return sendJson(res, 200, { session: publicSession(data.session) });
    }

    if (body.action === 'recovery') {
      const { error } = await supabase.auth.resetPasswordForEmail(body.email, {
        redirectTo: `${requestOrigin(req)}/?recovery=1`,
      });
      if (error) return sendJson(res, 400, { error: safeAuthError(error) });
      return sendJson(res, 200, { ok: true });
    }

    const { data, error } = await supabase.auth.setSession({
      access_token: body.accessToken,
      refresh_token: body.refreshToken,
    });
    if (error || !data.session) return sendJson(res, 401, { error: 'Invalid recovery session' });
    setAuthCookies(res, data.session);
    return sendJson(res, 200, { session: publicSession(data.session), recovery: true });
  } catch (error) {
    console.error('Auth API failure', error);
    return sendJson(res, 500, { error: 'Authentication service unavailable' });
  }
}
