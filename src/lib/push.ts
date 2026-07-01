export type PushNotificationType =
  | 'friend_invite'
  | 'friend_accepted'
  | 'shared_expense_created'
  | 'shared_income_created'
  | 'shared_approved'
  | 'shared_declined'
  | 'shopping_list_updated'
  | 'financial_alert';

export type PushPreferences = Record<PushNotificationType, boolean>;

export type PushStatus = {
  supported: boolean;
  permission: NotificationPermission | 'unsupported';
  subscribed: boolean;
  platform: 'ios' | 'android' | 'desktop' | 'unknown';
  standalone: boolean;
  publicKeyConfigured: boolean;
};

export const PUSH_PREFERENCE_LABELS: Array<{ key: PushNotificationType; label: string; description: string }> = [
  { key: 'friend_invite', label: 'Convites de amizade', description: 'Quando alguém enviar um convite para você.' },
  { key: 'friend_accepted', label: 'Aprovação de amizade', description: 'Quando aceitarem um convite enviado por você.' },
  { key: 'shared_expense_created', label: 'Despesas compartilhadas', description: 'Quando uma despesa precisar da sua aprovação.' },
  { key: 'shared_income_created', label: 'Receitas compartilhadas', description: 'Quando uma receita precisar da sua aprovação.' },
  { key: 'shopping_list_updated', label: 'Listas de compras', description: 'Quando uma lista compartilhada receber novos itens.' },
  { key: 'financial_alert', label: 'Alertas financeiros', description: 'Lembretes e avisos importantes da conta.' },
];

export const DEFAULT_PUSH_PREFERENCES: PushPreferences = {
  friend_invite: true,
  friend_accepted: true,
  shared_expense_created: true,
  shared_income_created: true,
  shared_approved: true,
  shared_declined: true,
  shopping_list_updated: true,
  financial_alert: true,
};

const PUBLIC_VAPID_KEY = import.meta.env.VITE_VAPID_PUBLIC_KEY || '';

function urlBase64ToUint8Array(value: string) {
  const padding = '='.repeat((4 - (value.length % 4)) % 4);
  const base64 = `${value}${padding}`.replace(/-/g, '+').replace(/_/g, '/');
  const raw = window.atob(base64);
  return Uint8Array.from([...raw].map((char) => char.charCodeAt(0)));
}

export function detectPushPlatform(): PushStatus['platform'] {
  const ua = navigator.userAgent.toLowerCase();
  if (/iphone|ipad|ipod/.test(ua)) return 'ios';
  if (/android/.test(ua)) return 'android';
  if (/windows|macintosh|linux|cros/.test(ua)) return 'desktop';
  return 'unknown';
}

export function isStandalonePwa() {
  return window.matchMedia('(display-mode: standalone)').matches || (navigator as Navigator & { standalone?: boolean }).standalone === true;
}

export function isPushSupported() {
  return 'serviceWorker' in navigator && 'PushManager' in window && 'Notification' in window;
}

async function pushRequest<T>(body: Record<string, unknown>): Promise<T> {
  const response = await fetch('/api/push', {
    method: 'POST',
    credentials: 'same-origin',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  const payload = await response.json().catch(() => ({})) as { error?: string } & T;
  if (!response.ok) throw new Error(payload.error || 'Não foi possível configurar notificações.');
  return payload;
}

export async function getPushStatus(): Promise<PushStatus> {
  const supported = isPushSupported();
  if (!supported) {
    return { supported: false, permission: 'unsupported', subscribed: false, platform: detectPushPlatform(), standalone: isStandalonePwa(), publicKeyConfigured: Boolean(PUBLIC_VAPID_KEY) };
  }
  const registration = await navigator.serviceWorker.ready;
  const subscription = await registration.pushManager.getSubscription();
  return { supported: true, permission: Notification.permission, subscribed: Boolean(subscription), platform: detectPushPlatform(), standalone: isStandalonePwa(), publicKeyConfigured: Boolean(PUBLIC_VAPID_KEY) };
}

export async function enablePushNotifications() {
  if (!isPushSupported()) throw new Error('Este navegador ou dispositivo não suporta notificações push para PWA.');
  if (!PUBLIC_VAPID_KEY) throw new Error('A chave pública VAPID não está configurada no RubyLife.');

  const permission = await Notification.requestPermission();
  if (permission === 'denied') throw new Error('Você recusou as notificações. Para receber avisos, será necessário ativar a permissão nas configurações do navegador ou do dispositivo.');
  if (permission !== 'granted') throw new Error('Permissão de notificações não concedida.');

  const registration = await navigator.serviceWorker.ready;
  const subscription = await registration.pushManager.subscribe({
    userVisibleOnly: true,
    applicationServerKey: urlBase64ToUint8Array(PUBLIC_VAPID_KEY),
  });

  await pushRequest<{ ok: true }>({
    action: 'register',
    subscription: subscription.toJSON(),
    platform: detectPushPlatform(),
    browser: navigator.userAgent.slice(0, 300),
  });
  return subscription;
}

export async function disablePushNotifications() {
  if (!isPushSupported()) return;
  const registration = await navigator.serviceWorker.ready;
  const subscription = await registration.pushManager.getSubscription();
  if (subscription) {
    await pushRequest<{ ok: true }>({ action: 'disable', endpoint: subscription.endpoint });
    await subscription.unsubscribe();
  }
}

export async function savePushPreferences(preferences: PushPreferences) {
  await pushRequest<{ ok: true }>({ action: 'preferences', preferences });
}

export async function sendPushNotification(input: {
  userId: string;
  title: string;
  body: string;
  url: string;
  type: PushNotificationType;
  data?: Record<string, unknown>;
}) {
  await pushRequest<{ ok: true }>({ action: 'send', ...input });
}


