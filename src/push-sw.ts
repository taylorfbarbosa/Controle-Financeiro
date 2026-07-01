/// <reference lib="webworker" />
import { cleanupOutdatedCaches, precacheAndRoute } from 'workbox-precaching';

declare let self: ServiceWorkerGlobalScope & { __WB_MANIFEST: Array<{ url: string; revision: string | null }> };

type PushPayload = {
  title?: string;
  body?: string;
  url?: string;
  type?: string;
  icon?: string;
  badge?: string;
  tag?: string;
  data?: Record<string, unknown>;
};

precacheAndRoute(self.__WB_MANIFEST);
cleanupOutdatedCaches();

const DEFAULT_ICON = '/icon-192.png';
const DEFAULT_BADGE = '/icon-192.png';

self.addEventListener('push', (event) => {
  const pushEvent = event as PushEvent;
  let payload: PushPayload = {};
  try {
    payload = pushEvent.data?.json() ?? {};
  } catch {
    payload = { body: pushEvent.data?.text() ?? '' };
  }

  const title = payload.title || 'RubyLife';
  const options: NotificationOptions = {
    body: payload.body || 'Você tem uma nova atualização no RubyLife.',
    icon: payload.icon || DEFAULT_ICON,
    badge: payload.badge || DEFAULT_BADGE,
    tag: payload.tag || payload.type || 'rubylife-update',
    data: {
      url: payload.url || '/',
      type: payload.type || 'general',
      ...(payload.data || {}),
    },
  };

  pushEvent.waitUntil(self.registration.showNotification(title, options));
});

self.addEventListener('notificationclick', (event) => {
  const notificationEvent = event as NotificationEvent;
  notificationEvent.notification.close();

  const rawUrl = String(notificationEvent.notification.data?.url || '/');
  const targetUrl = new URL(rawUrl, self.location.origin).href;

  notificationEvent.waitUntil((async () => {
    const clientList = await self.clients.matchAll({ type: 'window', includeUncontrolled: true });
    for (const client of clientList) {
      await client.focus();
      client.postMessage({ type: 'RUBYLIFE_PUSH_NAVIGATE', url: targetUrl });
      return;
    }
    await self.clients.openWindow(targetUrl);
  })());
});

