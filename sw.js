// CleanSwift Service Worker v1.0
const CACHE_NAME = 'cleanswift-v1';
const ASSETS = ['/', '/index.html', '/manifest.json', '/icon-192.png', '/icon-512.png'];

// ── INSTALLATION ──────────────────────────────────────────────────────────────
self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

// ── CACHE FIRST (offline support) ────────────────────────────────────────────
self.addEventListener('fetch', e => {
  e.respondWith(
    caches.match(e.request).then(cached => cached || fetch(e.request).catch(() => caches.match('/index.html')))
  );
});

// ── NOTIFICATIONS PUSH ────────────────────────────────────────────────────────
self.addEventListener('push', e => {
  const data = e.data ? e.data.json() : { title: 'CleanSwift', body: 'Rappel de service' };
  e.waitUntil(
    self.registration.showNotification(data.title, {
      body: data.body,
      icon: '/icon-192.png',
      badge: '/icon-192.png',
      tag: data.tag || 'cleanswift',
      data: data.url || '/',
      actions: [
        { action: 'view', title: '📋 Voir le client' },
        { action: 'dismiss', title: 'Ignorer' }
      ],
      vibrate: [200, 100, 200],
      requireInteraction: true
    })
  );
});

self.addEventListener('notificationclick', e => {
  e.notification.close();
  if (e.action === 'view' || !e.action) {
    e.waitUntil(clients.openWindow(e.notification.data || '/'));
  }
});

// ── ALARM SCHEDULER (vérification toutes les heures via sync) ─────────────────
// Les alarmes sont stockées dans IndexedDB et vérifiées au réveil du SW
self.addEventListener('periodicsync', e => {
  if (e.tag === 'check-reminders') {
    e.waitUntil(checkAndFireReminders());
  }
});

async function checkAndFireReminders() {
  // Lire les alarmes depuis le cache ou IDB
  const cache = await caches.open(CACHE_NAME);
  const resp = await cache.match('/reminders-data');
  if (!resp) return;
  const reminders = await resp.json();
  const now = Date.now();
  for (const r of reminders) {
    if (r.fireAt <= now && !r.fired) {
      await self.registration.showNotification(`🧹 CleanSwift — ${r.clientName}`, {
        body: r.message,
        icon: '/icon-192.png',
        tag: `reminder-${r.id}`,
        requireInteraction: true,
        vibrate: [300, 100, 300, 100, 300]
      });
    }
  }
}
