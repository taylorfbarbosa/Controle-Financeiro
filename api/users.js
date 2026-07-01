import {
  createServiceRoleClient,
  enforceRateLimit,
  getAuthenticatedContext,
  requestFingerprint,
  sendJson,
} from './_security.js';

function computeFriendId(userId) {
  const hex = userId.replace(/-/g, '');
  return (BigInt('0x' + hex) % 1000000n).toString().padStart(6, '0');
}

export default async function handler(req, res) {
  try {
    const context = await getAuthenticatedContext(req, res);
    if (!context) return sendJson(res, 401, { error: 'Unauthorized' });
    if (req.method !== 'GET') return sendJson(res, 405, { error: 'Method not allowed' });

    const rawInput = String(req.query?.friendId ?? '').trim().replace(/^#/, '');
    if (!rawInput) return sendJson(res, 400, { error: 'Invalid friend ID' });

    // Accept 6-digit numeric ID or UUID
    const isUuid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(rawInput);
    const numeric6 = rawInput.replace(/\D/g, '').padStart(6, '0').slice(-6);
    const validNumeric = /^\d{6}$/.test(numeric6) && rawInput.replace(/\D/g, '').length <= 6;

    if (!isUuid && !validNumeric) return sendJson(res, 400, { error: 'Invalid friend ID format' });

    const allowed = await enforceRateLimit(requestFingerprint(req, 'users-search', context.user.id), 30, 60);
    if (!allowed) return sendJson(res, 429, { error: 'Too many requests. Try again later.' });

    // Use service role if available, fall back to authenticated client
    // With the profiles_select_authenticated policy the regular client works too
    let readClient = context.client;
    let writeClient = null;
    try {
      const admin = createServiceRoleClient();
      readClient = admin;
      writeClient = admin;
    } catch (_) {
      // No service role key — authenticated client reads fine with the new policy
    }

    let foundProfile = null;

    // ── Strategy 1: RPC function (SECURITY DEFINER, works regardless of RLS) ──
    if (validNumeric && !foundProfile) {
      try {
        const { data } = await context.client.rpc('find_profile_by_friend_id', { p_friend_id: numeric6 });
        if (data && data.length > 0) foundProfile = data[0];
      } catch (_) {}
    }

    // ── Strategy 2: Direct query on friend_id column ──
    if (validNumeric && !foundProfile) {
      try {
        const { data } = await readClient
          .from('profiles')
          .select('id, full_name, avatar_url, friend_id')
          .eq('friend_id', numeric6)
          .maybeSingle();
        if (data) foundProfile = data;
      } catch (_) {}
    }

    // ── Strategy 3: UUID direct lookup ──
    if (isUuid && !foundProfile) {
      try {
        const { data } = await readClient
          .from('profiles')
          .select('id, full_name, avatar_url, friend_id')
          .eq('id', rawInput)
          .maybeSingle();
        if (data) foundProfile = data;
      } catch (_) {}
    }

    // ── Strategy 4: Full scan fallback (compute friend_id from uuid in-memory) ──
    if (!foundProfile && validNumeric) {
      try {
        const { data: allProfiles } = await readClient
          .from('profiles')
          .select('id, full_name, avatar_url, friend_id');

        if (allProfiles?.length > 0) {
          const match = allProfiles.find((p) =>
            p.friend_id === numeric6 || computeFriendId(p.id) === numeric6,
          );
          if (match) foundProfile = match;
        }
      } catch (_) {}
    }

    if (!foundProfile) return sendJson(res, 200, { user: null });

    const publicId = foundProfile.friend_id || computeFriendId(foundProfile.id);

    // Backfill friend_id if missing (fire-and-forget, needs write permission)
    if (!foundProfile.friend_id) {
      const wc = writeClient ?? context.client;
      wc.from('profiles').update({ friend_id: publicId }).eq('id', foundProfile.id)
        .then(() => {}).catch(() => {});
    }

    return sendJson(res, 200, {
      user: {
        id: foundProfile.id,
        publicFriendId: publicId,
        name: foundProfile.full_name || 'Usuário',
        email: '',
        avatarUrl: foundProfile.avatar_url ?? null,
      },
    });
  } catch (err) {
    console.error('[api/users] error:', err);
    return sendJson(res, 500, { error: 'Internal server error' });
  }
}
