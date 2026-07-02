import { z } from 'zod';
import {
  assertSameOrigin,
  enforceRateLimit,
  getAuthenticatedContext,
  requestFingerprint,
  sendJson,
} from './_security.js';

const uuid = z.string().uuid();

function mapFriendship(row) {
  return {
    id: row.id,
    requesterId: row.requester_id,
    receiverId: row.receiver_id,
    status: row.status,
    createdAt: row.created_at,
    updatedAt: row.updated_at,
  };
}

function mapProfile(row) {
  if (!row) return null;
  return {
    id: row.id,
    name: row.full_name || row.email || 'Usuário',
    email: row.email || '',
    avatarUrl: row.avatar_url || null,
    publicFriendId: row.friend_id || null,
  };
}

export default async function handler(req, res) {
  try {
    const context = await getAuthenticatedContext(req, res);
    if (!context) return sendJson(res, 401, { error: 'Unauthorized' });

    const uid = context.user.id;
    const client = context.client;

    // ────────────── GET: list friendships + profiles ──────────────
    if (req.method === 'GET') {
      const { data: friendships, error: fsErr } = await client
        .from('friendships')
        .select('*')
        .or(`requester_id.eq.${uid},receiver_id.eq.${uid}`)
        .order('created_at', { ascending: false });
      if (fsErr) throw fsErr;

      // Collect peer IDs
      const peerIds = [...new Set(
        (friendships || []).map((f) => f.requester_id === uid ? f.receiver_id : f.requester_id)
      )];

      let profiles = [];
      if (peerIds.length) {
        const { data: profileRows, error: prErr } = await client
          .from('profiles')
          .select('id, full_name, email, avatar_url, friend_id')
          .in('id', peerIds);
        if (prErr) throw prErr;
        profiles = (profileRows || []).map(mapProfile);
      }

      return sendJson(res, 200, {
        friendships: (friendships || []).map(mapFriendship),
        profiles,
      });
    }

    if (!assertSameOrigin(req)) return sendJson(res, 403, { error: 'Invalid request origin' });
    const allowed = await enforceRateLimit(requestFingerprint(req, 'friends', uid), 60, 60);
    if (!allowed) return sendJson(res, 429, { error: 'Too many requests.' });

    // ────────────── POST: send request or lookup ──────────────
    if (req.method === 'POST') {
      const body = req.body || {};

      // Lookup by friend_id
      if (body.action === 'lookup') {
        const parsed = z.object({ friendId: z.string().min(1).max(20) }).safeParse(body);
        if (!parsed.success) return sendJson(res, 400, { error: 'friendId inválido' });

        const { data: rows, error } = await client.rpc('find_profile_by_friend_id', {
          p_friend_id: parsed.data.friendId,
        });
        if (error) throw error;
        const profile = rows && rows.length ? mapProfile({ ...rows[0], email: rows[0].email || '' }) : null;
        return sendJson(res, 200, { profile });
      }

      // Send friend request
      const parsed = z.object({ receiverId: uuid }).safeParse(body);
      if (!parsed.success) return sendJson(res, 400, { error: 'receiverId inválido' });
      const { receiverId } = parsed.data;
      if (receiverId === uid) return sendJson(res, 400, { error: 'Você não pode se adicionar como amigo.' });

      // Check for existing friendship
      const { data: existing } = await client
        .from('friendships')
        .select('id, status')
        .or(`and(requester_id.eq.${uid},receiver_id.eq.${receiverId}),and(requester_id.eq.${receiverId},receiver_id.eq.${uid})`)
        .maybeSingle();

      if (existing) {
        if (existing.status === 'accepted') return sendJson(res, 409, { error: 'Já são amigos.' });
        if (existing.status === 'pending') return sendJson(res, 409, { error: 'Já existe um convite pendente.' });
      }

      const { data: inserted, error: insErr } = await client
        .from('friendships')
        .insert({ requester_id: uid, receiver_id: receiverId, status: 'pending' })
        .select()
        .single();
      if (insErr) throw insErr;

      return sendJson(res, 201, { friendship: mapFriendship(inserted) });
    }

    // ────────────── PATCH: accept or decline ──────────────
    if (req.method === 'PATCH') {
      const parsed = z.object({
        id: uuid,
        status: z.enum(['accepted', 'declined']),
      }).safeParse(req.body);
      if (!parsed.success) return sendJson(res, 400, { error: 'Dados inválidos' });

      const { data: current } = await client
        .from('friendships')
        .select('*')
        .eq('id', parsed.data.id)
        .maybeSingle();

      if (!current) return sendJson(res, 404, { error: 'Convite não encontrado' });
      if (current.receiver_id !== uid) return sendJson(res, 403, { error: 'Sem permissão para responder este convite.' });
      if (current.status !== 'pending') return sendJson(res, 409, { error: 'Convite não está pendente.' });

      const { data: updated, error: updErr } = await client
        .from('friendships')
        .update({ status: parsed.data.status })
        .eq('id', parsed.data.id)
        .select()
        .single();
      if (updErr) throw updErr;

      return sendJson(res, 200, { friendship: mapFriendship(updated) });
    }

    // ────────────── DELETE: remove friendship ──────────────
    if (req.method === 'DELETE') {
      const parsed = z.object({ id: uuid }).safeParse(req.body);
      if (!parsed.success) return sendJson(res, 400, { error: 'id inválido' });

      const { error: delErr } = await client
        .from('friendships')
        .delete()
        .eq('id', parsed.data.id)
        .or(`requester_id.eq.${uid},receiver_id.eq.${uid}`);
      if (delErr) throw delErr;

      return sendJson(res, 200, { ok: true });
    }

    return sendJson(res, 405, { error: 'Method not allowed' });
  } catch (err) {
    console.error('Friends API failure', err);
    return sendJson(res, 500, { error: 'Erro interno ao processar solicitação de amizade.' });
  }
}
