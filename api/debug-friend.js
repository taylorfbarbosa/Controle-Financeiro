/**
 * Endpoint de diagnóstico temporário — remova após resolver o problema.
 * GET /api/debug-friend?friendId=XXXXXX
 */
import {
  createServiceRoleClient,
  getAuthenticatedContext,
  sendJson,
} from './_security.js';

function computeFriendId(userId) {
  const hex = userId.replace(/-/g, '');
  return (BigInt('0x' + hex) % 1000000n).toString().padStart(6, '0');
}

export default async function handler(req, res) {
  const context = await getAuthenticatedContext(req, res);
  if (!context) return sendJson(res, 401, { error: 'Unauthorized' });

  const friendId = String(req.query?.friendId ?? '').trim();
  const report = { friendId, steps: [] };

  // Step 1: Check if service role is available
  let adminClient = null;
  try {
    adminClient = createServiceRoleClient();
    report.steps.push({ step: 'service_role', ok: true });
  } catch (e) {
    report.steps.push({ step: 'service_role', ok: false, error: e.message });
  }

  const readClient = adminClient ?? context.client;

  // Step 2: Check if friend_id column exists
  try {
    const { data, error } = await readClient
      .from('profiles')
      .select('friend_id')
      .limit(1);
    report.steps.push({
      step: 'column_friend_id_exists',
      ok: !error,
      error: error?.message,
      sample: data?.[0] ?? null,
    });
  } catch (e) {
    report.steps.push({ step: 'column_friend_id_exists', ok: false, error: e.message });
  }

  // Step 3: Count profiles readable
  try {
    const { count, error } = await readClient
      .from('profiles')
      .select('id', { count: 'exact', head: true });
    report.steps.push({
      step: 'profiles_readable_count',
      ok: !error,
      count: count ?? 0,
      error: error?.message,
    });
  } catch (e) {
    report.steps.push({ step: 'profiles_readable_count', ok: false, error: e.message });
  }

  // Step 4: Try RPC function
  try {
    const { data, error } = await context.client.rpc('find_profile_by_friend_id', { p_friend_id: friendId });
    report.steps.push({
      step: 'rpc_find_profile_by_friend_id',
      ok: !error,
      found: data?.length > 0,
      data: data?.[0] ?? null,
      error: error?.message,
    });
  } catch (e) {
    report.steps.push({ step: 'rpc_find_profile_by_friend_id', ok: false, error: e.message });
  }

  // Step 5: Direct query by friend_id
  if (friendId) {
    try {
      const { data, error } = await readClient
        .from('profiles')
        .select('id, full_name, friend_id')
        .eq('friend_id', friendId)
        .maybeSingle();
      report.steps.push({
        step: 'direct_query_by_friend_id',
        ok: !error,
        found: !!data,
        data: data ?? null,
        error: error?.message,
      });
    } catch (e) {
      report.steps.push({ step: 'direct_query_by_friend_id', ok: false, error: e.message });
    }
  }

  // Step 6: Show own profile with computed friend_id
  try {
    const { data } = await context.client
      .from('profiles')
      .select('id, full_name, friend_id')
      .eq('id', context.user.id)
      .maybeSingle();
    report.steps.push({
      step: 'own_profile',
      ok: true,
      profile: data,
      computedFriendId: computeFriendId(context.user.id),
    });
  } catch (e) {
    report.steps.push({ step: 'own_profile', ok: false, error: e.message });
  }

  return sendJson(res, 200, report);
}
