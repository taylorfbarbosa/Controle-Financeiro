import webpush from 'web-push';
import { z } from 'zod';
import { createClient } from '@supabase/supabase-js';
import {
  assertSameOrigin,
  enforceRateLimit,
  getAuthenticatedContext,
  requestFingerprint,
  sendJson,
  supabaseConfig,
} from './_security.js';

const pushTypes = z.enum([
  'friend_invite',
  'friend_accepted',
  'shared_expense_created',
  'shared_income_created',
  'shared_approved',
  'shared_declined',
  'shopping_list_updated',
  'financial_alert',
]);

const subscriptionSchema = z.object({
  endpoint: z.string().url().max(2000),
  keys: z.object({
    p256dh: z.string().min(20).max(500),
    auth: z.string().min(10).max(300),
  }).strict(),
}).passthrough();

const registerSchema = z.object({
  action: z.literal('register'),
  subscription: subscriptionSchema,
  platform: z.string().trim().max(40).optional(),
  browser: z.string().trim().max(320).optional(),
}).strict();

const disableSchema = z.object({
  action: z.literal('disable'),
  endpoint: z.string().url().max(2000),
}).strict();

const preferencesSchema = z.object({
  action: z.literal('preferences'),
  preferences: z.record(pushTypes, z.boolean()).partial(),
}).strict();

const sendSchema = z.object({
  action: z.literal('send'),
  userId: z.string().uuid().or(z.string().min(3).max(120)),
  title: z.string().trim().min(1).max(80),
  body: z.string().trim().min(1).max(180),
  url: z.string().trim().min(1).max(600),
  type: pushTypes,
  data: z.record(z.string(), z.unknown()).optional(),
}).strict();

function adminClient() {
  const serviceRoleKey = process.env.SUPABASE_SERVICE_ROLE_KEY;
  if (!serviceRoleKey) return null;
  const { url } = supabaseConfig();
  return createClient(url, serviceRoleKey, {
    auth: { persistSession: false, autoRefreshToken: false, detectSessionInUrl: false },
  });
}

function configureWebPush() {
  const publicKey = process.env.VAPID_PUBLIC_KEY || process.env.VITE_VAPID_PUBLIC_KEY;
  const privateKey = process.env.VAPID_PRIVATE_KEY;
  const subject = process.env.VAPID_SUBJECT || 'mailto:suporte@rubylife.app';
  if (!publicKey || !privateKey) throw new Error('VAPID keys are not configured');
  webpush.setVapidDetails(subject, publicKey, privateKey);
}

function sameOriginUrl(req, url) {
  try {
    const origin = `${req.headers['x-forwarded-proto'] || 'https'}://${req.headers['x-forwarded-host'] || req.headers.host}`;
    const parsed = new URL(url, origin);
    if (parsed.origin !== origin) return '/';
    return `${parsed.pathname}${parsed.search}${parsed.hash}`;
  } catch {
    return '/';
  }
}

async function createInternalHistory(client, payload) {
  await client.from('notification_history').insert({
    user_id: payload.userId,
    title: payload.title,
    message: payload.body,
    type: payload.type,
    target_url: payload.url,
    read: false,
  });
}

async function sendToUser(req, currentUserId, payload) {
  configureWebPush();
  const admin = adminClient();
  if (!admin) throw new Error('SUPABASE_SERVICE_ROLE_KEY is required to send push notifications');

  const { data: preferences } = await admin
    .from('notification_preferences')
    .select('preferences')
    .eq('user_id', payload.userId)
    .maybeSingle();
  if (preferences?.preferences?.[payload.type] === false) {
    await createInternalHistory(admin, payload).catch(() => undefined);
    return { sent: 0, skippedByPreference: true };
  }

  const { data: devices, error } = await admin
    .from('push_devices')
    .select('id, endpoint, p256dh_key, auth_key')
    .eq('user_id', payload.userId)
    .eq('active', true);
  if (error) throw error;

  const notification = JSON.stringify({
    title: payload.title,
    body: payload.body,
    url: sameOriginUrl(req, payload.url),
    type: payload.type,
    icon: '/icon-192.png',
    badge: '/icon-192.png',
    data: { ...(payload.data || {}), sentBy: currentUserId },
  });

  let sent = 0;
  for (const device of devices || []) {
    try {
      await webpush.sendNotification({
        endpoint: device.endpoint,
        keys: { p256dh: device.p256dh_key, auth: device.auth_key },
      }, notification);
      sent += 1;
      await admin.from('push_devices').update({ last_used_at: new Date().toISOString() }).eq('id', device.id);
    } catch (error) {
      const statusCode = error?.statusCode || error?.status;
      if (statusCode === 404 || statusCode === 410) {
        await admin.from('push_devices').update({ active: false }).eq('id', device.id);
      } else {
        console.error('Push send failure', error);
      }
    }
  }

  await createInternalHistory(admin, payload).catch((error) => console.error('Notification history failure', error));
  return { sent };
}

export default async function handler(req, res) {
  try {
    const context = await getAuthenticatedContext(req, res);
    if (!context) return sendJson(res, 401, { error: 'Unauthorized' });
    if (req.method !== 'POST') return sendJson(res, 405, { error: 'Method not allowed' });
    if (!assertSameOrigin(req)) return sendJson(res, 403, { error: 'Invalid request origin' });

    const allowed = await enforceRateLimit(requestFingerprint(req, 'push', context.user.id), 120, 60);
    if (!allowed) return sendJson(res, 429, { error: 'Too many requests. Try again later.' });

    const register = registerSchema.safeParse(req.body);
    if (register.success) {
      const row = {
        user_id: context.user.id,
        endpoint: register.data.subscription.endpoint,
        p256dh_key: register.data.subscription.keys.p256dh,
        auth_key: register.data.subscription.keys.auth,
        platform: register.data.platform || null,
        browser: register.data.browser || null,
        active: true,
        last_used_at: new Date().toISOString(),
      };
      const { error } = await context.client.from('push_devices').upsert(row, { onConflict: 'endpoint' });
      if (error) throw error;
      return sendJson(res, 200, { ok: true });
    }

    const disable = disableSchema.safeParse(req.body);
    if (disable.success) {
      const { error } = await context.client
        .from('push_devices')
        .update({ active: false })
        .eq('user_id', context.user.id)
        .eq('endpoint', disable.data.endpoint);
      if (error) throw error;
      return sendJson(res, 200, { ok: true });
    }

    const preferences = preferencesSchema.safeParse(req.body);
    if (preferences.success) {
      const { error } = await context.client.from('notification_preferences').upsert({
        user_id: context.user.id,
        preferences: preferences.data.preferences,
        updated_at: new Date().toISOString(),
      }, { onConflict: 'user_id' });
      if (error) throw error;
      return sendJson(res, 200, { ok: true });
    }

    const send = sendSchema.safeParse(req.body);
    if (send.success) {
      const result = await sendToUser(req, context.user.id, send.data);
      return sendJson(res, 200, { ok: true, ...result });
    }

    return sendJson(res, 400, { error: 'Invalid push request' });
  } catch (error) {
    console.error('Push API failure', error);
    const message = error instanceof Error && error.message.includes('VAPID')
      ? 'Push notifications are not configured on the server.'
      : 'Unable to process push request';
    return sendJson(res, 500, { error: message });
  }
}
