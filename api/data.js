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
const shortText = z.string().trim().min(1).max(160);
const optionalText = z.string().trim().max(2000).nullable();
const date = z.string().regex(/^\d{4}-\d{2}-\d{2}$/);
const color = z.string().regex(/^#[0-9a-f]{6}$/i).nullable();

const schemas = {
  accounts: z.object({
    id: uuid, name: shortText, type: z.enum(['wallet', 'checking', 'savings', 'investment']),
    initial_balance: z.number().finite().min(-1e12).max(1e12), color, icon: z.string().max(50).nullable(),
  }).strict(),
  categories: z.object({
    id: uuid, name: shortText, kind: z.enum(['income', 'expense', 'transfer']), color, icon: z.string().max(50).nullable(),
  }).strict(),
  transactions: z.object({
    id: uuid, group_id: uuid, description: z.string().trim().min(1).max(200), category: shortText,
    category_id: nullableUuid, amount: z.number().finite().min(0).max(1e12), type: z.enum(['income', 'expense']),
    due_date: date, recurrence: z.enum(['single', 'fixed', 'installment']), installment_number: z.number().int().min(1).max(120).nullable(),
    installment_total: z.number().int().min(1).max(120).nullable(), status: z.enum(['open', 'settled']), settled_at: date.nullable(),
    settled_amount: z.number().finite().min(0).max(1e12).nullable(), account_id: nullableUuid, account: z.string().trim().max(160).nullable(), notes: optionalText,
  }).strict(),
  goals: z.object({
    id: uuid, name: shortText, target_amount: z.number().finite().positive().max(1e12), deadline: date.nullable(),
    color: z.string().regex(/^#[0-9a-f]{6}$/i), icon: z.string().max(50), created_at: date,
  }).strict(),
  goal_movements: z.object({
    id: uuid, goal_id: uuid, type: z.enum(['deposit', 'withdraw']), amount: z.number().finite().positive().max(1e12),
    date, note: optionalText,
  }).strict(),
};

const syncSchema = z.object({
  resource: z.enum(['accounts', 'categories', 'transactions', 'goals']),
  upsert: z.array(z.unknown()).max(1000),
  deleteIds: z.array(uuid).max(1000),
  replaceMovements: z.array(z.object({ goalId: uuid, rows: z.array(z.unknown()).max(1000) })).max(1000).optional(),
}).strict();

const profileUpdateSchema = z.object({
  resource: z.literal('profile'),
  profile: z.object({
    fullName: z.string().trim().min(2).max(160),
    phone: z.string().trim().max(30).regex(/^[0-9()+\-\s]*$/),
    avatarUrl: z.string().max(800000).nullable().optional(),
  }).strict(),
}).strict();

function isMissingProfilePhone(error) {
  if (!error) return false;
  const message = String(error.message || error.details || '').toLowerCase();
  return (error.code === '42703' || error.code === 'PGRST204') && message.includes('phone');
}

async function loadProfile(client) {
  const profile = await client.from('profiles').select('full_name, email, phone, avatar_url').maybeSingle();
  if (!isMissingProfilePhone(profile.error)) return profile;

  const fallback = await client.from('profiles').select('full_name, email, avatar_url').maybeSingle();
  return {
    ...fallback,
    data: fallback.data ? { ...fallback.data, phone: null } : null,
  };
}

async function loadAll(client) {
  const [profile, accounts, categories, transactions, goals, movements] = await Promise.all([
    loadProfile(client),
    client.from('accounts').select('*').order('created_at', { ascending: true }),
    client.from('categories').select('*').order('created_at', { ascending: true }),
    client.from('transactions').select('*').order('due_date', { ascending: true }),
    client.from('goals').select('*').order('created_at', { ascending: true }),
    client.from('goal_movements').select('*').order('date', { ascending: true }),
  ]);
  const error = profile.error || accounts.error || categories.error || transactions.error || goals.error || movements.error;
  if (error) throw error;
  return {
    profile: profile.data || null,
    accounts: accounts.data || [], categories: categories.data || [], transactions: transactions.data || [],
    goals: goals.data || [], goal_movements: movements.data || [],
  };
}

export default async function handler(req, res) {
  try {
    const context = await getAuthenticatedContext(req, res);
    if (!context) return sendJson(res, 401, { error: 'Unauthorized' });

    if (req.method === 'GET') return sendJson(res, 200, await loadAll(context.client));
    if (req.method !== 'POST') return sendJson(res, 405, { error: 'Method not allowed' });
    if (!assertSameOrigin(req)) return sendJson(res, 403, { error: 'Invalid request origin' });

    const allowed = await enforceRateLimit(requestFingerprint(req, 'data-write', context.user.id), 120, 60);
    if (!allowed) return sendJson(res, 429, { error: 'Too many requests. Try again later.' });

    const profileUpdate = profileUpdateSchema.safeParse(req.body);
    if (profileUpdate.success) {
      const updatePayload = {
        full_name: profileUpdate.data.profile.fullName,
        phone: profileUpdate.data.profile.phone || null,
      };
      if (profileUpdate.data.profile.avatarUrl !== undefined) {
        updatePayload.avatar_url = profileUpdate.data.profile.avatarUrl || null;
      }

      let result = await context.client
        .from('profiles')
        .update(updatePayload)
        .eq('id', context.user.id)
        .select('full_name, email, phone, avatar_url')
        .maybeSingle();

      if (isMissingProfilePhone(result.error)) {
        const { phone: _phone, ...fallbackPayload } = updatePayload;
        result = await context.client
          .from('profiles')
          .update(fallbackPayload)
          .eq('id', context.user.id)
          .select('full_name, email, avatar_url')
          .maybeSingle();
        if (result.data) result.data = { ...result.data, phone: null };
      }

      const { data, error } = result;
      if (error) throw error;
      if (!data) return sendJson(res, 404, { error: 'User profile not found' });
      return sendJson(res, 200, { profile: data });
    }

    const parsed = syncSchema.safeParse(req.body);
    if (!parsed.success) return sendJson(res, 400, { error: 'Invalid request data' });
    const { resource, deleteIds, replaceMovements } = parsed.data;
    const validatedRows = z.array(schemas[resource]).max(1000).safeParse(parsed.data.upsert);
    if (!validatedRows.success) return sendJson(res, 400, { error: 'Invalid resource data' });
    const rows = validatedRows.data.map((row) => ({ ...row, user_id: context.user.id }));

    if (deleteIds.length) {
      const { error } = await context.client.from(resource).delete().in('id', deleteIds);
      if (error) throw error;
    }
    if (rows.length) {
      const { error } = await context.client.from(resource).upsert(rows);
      if (error) throw error;
    }

    if (resource === 'goals' && replaceMovements) {
      for (const replacement of replaceMovements) {
        const movementRows = z.array(schemas.goal_movements).max(1000).safeParse(replacement.rows);
        if (!movementRows.success || movementRows.data.some((row) => row.goal_id !== replacement.goalId)) {
          return sendJson(res, 400, { error: 'Invalid goal movement data' });
        }
        const { error: deleteError } = await context.client.from('goal_movements').delete().eq('goal_id', replacement.goalId);
        if (deleteError) throw deleteError;
        if (movementRows.data.length) {
          const payload = movementRows.data.map((row) => ({ ...row, user_id: context.user.id }));
          const { error: insertError } = await context.client.from('goal_movements').insert(payload);
          if (insertError) throw insertError;
        }
      }
    }

    return sendJson(res, 200, { ok: true });
  } catch (error) {
    console.error('Data API failure', error);
    return sendJson(res, 500, { error: 'Unable to process data request' });
  }
}
