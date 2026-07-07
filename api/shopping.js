import { z } from 'zod';
import {
  assertSameOrigin,
  enforceRateLimit,
  getAuthenticatedContext,
  requestFingerprint,
  sendJson,
} from './_security.js';

const uuid = z.string().uuid();
const nullableUuid = uuid.nullable();
const shortText = z.string().trim().min(1).max(200);
const optionalText = z.string().trim().max(2000).nullable().optional();
const dateStr = z.string().regex(/^\d{4}-\d{2}-\d{2}$/).nullable().optional();
const isoStr = z.string().max(100).nullable().optional();
const statusEnum = z.enum(['open', 'finalized', 'cancelled']);

const itemSchema = z.object({
  id: uuid,
  name: z.string().trim().min(1).max(200),
  quantity: z.number().finite().positive().max(10000),
  unitPrice: z.number().finite().min(0).max(1e12),
  addedById: uuid,
  createdAt: z.string().max(100),
  purchased: z.boolean().optional(),
});

const listSchema = z.object({
  id: uuid,
  name: shortText,
  date: dateStr,
  note: optionalText,
  status: statusEnum,
  creatorId: uuid,
  participantIds: z.array(uuid).min(1).max(50),
  items: z.array(itemSchema).max(500),
  createdAt: z.string().max(100),
  finalizedAt: isoStr,
  finalizedById: nullableUuid.optional(),
  cancelledAt: isoStr,
});

function toRow(list, userId) {
  return {
    id: list.id,
    user_id: userId,
    name: list.name,
    date: list.date || null,
    note: list.note || null,
    status: list.status,
    creator_id: list.creatorId,
    participant_ids: list.participantIds,
    items: list.items,
    created_at: list.createdAt,
    finalized_at: list.finalizedAt || null,
    finalized_by_id: list.finalizedById || null,
    cancelled_at: list.cancelledAt || null,
  };
}

function fromRow(row) {
  return {
    id: row.id,
    name: row.name,
    date: row.date || '',
    note: row.note || undefined,
    status: row.status,
    creatorId: row.creator_id,
    participantIds: row.participant_ids || [],
    items: row.items || [],
    createdAt: row.created_at,
    finalizedAt: row.finalized_at || undefined,
    finalizedById: row.finalized_by_id || undefined,
    cancelledAt: row.cancelled_at || undefined,
  };
}

export default async function handler(req, res) {
  try {
    const context = await getAuthenticatedContext(req, res);
    if (!context) return sendJson(res, 401, { error: 'Unauthorized' });

    const uid = context.user.id;
    const client = context.client;

    // ────────────── GET: list all lists for this user ──────────────
    if (req.method === 'GET') {
      const { data, error } = await client
        .from('shopping_lists')
        .select('*')
        .order('created_at', { ascending: false });
      if (error) throw error;
      return sendJson(res, 200, { lists: (data || []).map(fromRow) });
    }

    if (!assertSameOrigin(req)) return sendJson(res, 403, { error: 'Invalid request origin' });
    const allowed = await enforceRateLimit(requestFingerprint(req, 'shopping', uid), 120, 60);
    if (!allowed) return sendJson(res, 429, { error: 'Too many requests.' });

    // ────────────── POST: create list ──────────────
    if (req.method === 'POST') {
      const parsed = listSchema.safeParse(req.body);
      if (!parsed.success) return sendJson(res, 400, { error: 'Dados inválidos', details: parsed.error.issues });

      // Ensure current user is the creator and a participant
      if (parsed.data.creatorId !== uid) return sendJson(res, 403, { error: 'creatorId deve ser o usuário atual.' });
      if (!parsed.data.participantIds.includes(uid)) return sendJson(res, 400, { error: 'Criador deve ser participante.' });

      const row = toRow(parsed.data, uid);
      const { data, error } = await client.from('shopping_lists').insert(row).select().single();
      if (error) throw error;
      return sendJson(res, 201, { list: fromRow(data) });
    }

    // ────────────── PATCH: update list ──────────────
    if (req.method === 'PATCH') {
      const parsed = listSchema.safeParse(req.body);
      if (!parsed.success) return sendJson(res, 400, { error: 'Dados inválidos' });

      // Verify list exists and user is creator or participant
      const { data: current, error: fetchErr } = await client
        .from('shopping_lists')
        .select('creator_id, participant_ids, status')
        .eq('id', parsed.data.id)
        .maybeSingle();
      if (fetchErr) throw fetchErr;
      if (!current) return sendJson(res, 404, { error: 'Lista não encontrada.' });

      const isCreator = current.creator_id === uid;
      const isParticipant = (current.participant_ids || []).includes(uid);
      if (!isCreator && !isParticipant) return sendJson(res, 403, { error: 'Sem permissão para atualizar esta lista.' });

      // Only creator can change status to cancelled or change participants
      const newStatus = parsed.data.status;
      if (newStatus === 'cancelled' && !isCreator) return sendJson(res, 403, { error: 'Somente o criador pode cancelar.' });

      const row = toRow(parsed.data, current.creator_id === uid ? uid : current.creator_id);
      const { data, error } = await client
        .from('shopping_lists')
        .update(row)
        .eq('id', parsed.data.id)
        .select()
        .single();
      if (error) throw error;
      return sendJson(res, 200, { list: fromRow(data) });
    }

    // ────────────── DELETE: delete list ──────────────
    if (req.method === 'DELETE') {
      const parsed = z.object({ id: uuid }).safeParse(req.body);
      if (!parsed.success) return sendJson(res, 400, { error: 'id inválido' });

      // Only creator (= user_id) can delete
      const { error: delErr } = await client
        .from('shopping_lists')
        .delete()
        .eq('id', parsed.data.id)
        .eq('user_id', uid);
      if (delErr) throw delErr;

      return sendJson(res, 200, { ok: true });
    }

    return sendJson(res, 405, { error: 'Method not allowed' });
  } catch (err) {
    console.error('Shopping API failure', err);
    return sendJson(res, 500, { error: 'Erro interno ao processar lista de compras.' });
  }
}
