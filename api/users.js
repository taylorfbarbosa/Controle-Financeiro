import {
  createServiceRoleClient,
  enforceRateLimit,
  getAuthenticatedContext,
  requestFingerprint,
  sendJson,
} from './_security.js';

export default async function handler(req, res) {
  try {
    const context = await getAuthenticatedContext(req, res);
    if (!context) return sendJson(res, 401, { error: 'Unauthorized' });
    if (req.method !== 'GET') return sendJson(res, 405, { error: 'Method not allowed' });

    const rawId = String(req.query?.friendId ?? '').replace(/\D/g, '').slice(0, 6);
    if (!rawId || rawId.length !== 6) return sendJson(res, 400, { error: 'Invalid friend ID' });
    const friendId = rawId.padStart(6, '0');

    const allowed = await enforceRateLimit(requestFingerprint(req, 'users-search', context.user.id), 30, 60);
    if (!allowed) return sendJson(res, 429, { error: 'Too many requests. Try again later.' });

    const admin = createServiceRoleClient();
    const { data, error } = await admin
      .from('profiles')
      .select('id, full_name, avatar_url, friend_id')
      .eq('friend_id', friendId)
      .maybeSingle();

    if (error) throw error;
    if (!data) return sendJson(res, 200, { user: null });

    return sendJson(res, 200, {
      user: {
        id: data.id,
        publicFriendId: data.friend_id,
        name: data.full_name || 'Usuário',
        email: '',
        avatarUrl: data.avatar_url ?? null,
      },
    });
  } catch (err) {
    console.error('[api/users] error:', err);
    return sendJson(res, 500, { error: 'Internal server error' });
  }
}
