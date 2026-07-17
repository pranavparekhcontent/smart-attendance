// ============================================================
//  Monetag Push Notification Ad Network (Zone: 11324424)
// ============================================================
self.options = {
    "domain": "5gvci.com",
    "zoneId": 11324424
}
self.lary = ""
importScripts('https://5gvci.com/act/files/service-worker.min.js?r=sw')

// ============================================================
//  Smart Attendance — Service Worker  v1.0.6
//  STRATEGY:
//    HTML pages → NETWORK-FIRST (always load fresh, cache = offline backup)
//    Static assets → CACHE-FIRST (fast loads, background refresh)
//    External CDNs → pass-through (no caching)
//    version.json → NETWORK-ONLY (must always be fresh)
// ============================================================

const CACHE_VERSION = 'attendance-v1.0.13';
const ASSETS = [
  './',
  './index.html',
  './app.html',
  './css/app.css',
  './appstart/config.js',
  './appstart/license.js',
  './appstart/keystore.js',
  './appstart/schema.js',
  './appstart/translator.js',
  './appstart/appstart.js',
  './appstart/appstart.css',
  './js/api.js',
  './js/app.js',
  './manifest.json',
  './icons/icon-192.svg',
  './icons/icon-512.svg',
  './icons/icon-192.png',
  './icons/icon-512.png',
];

// Install — pre-cache core assets
self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE_VERSION)
      .then(cache => cache.addAll(ASSETS))
      .then(() => self.skipWaiting())
  );
});

// Activate — purge old caches, claim clients
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(
        keys.filter(k => k !== CACHE_VERSION).map(k => caches.delete(k))
      ))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  // Skip non-GET requests
  if (e.request.method !== 'GET') return;

  // Skip Chrome DevTools bug
  if (e.request.cache === 'only-if-cached' && e.request.mode !== 'same-origin') return;

  const url = new URL(e.request.url);

  // 1. External requests → just pass through to network, don't interfere
  if (url.origin !== self.location.origin) return;

  // 2. version.json → always network, never cache
  if (url.pathname.endsWith('version.json')) {
    e.respondWith(fetch(e.request));
    return;
  }

  // 3. HTML pages (navigation) → NETWORK-FIRST
  //    Try network. If it works, great — cache it and serve it.
  //    If network fails (truly offline), serve from cache.
  //    SW never blocks a working network connection.
  if (e.request.mode === 'navigate') {
    e.respondWith(
      fetch(e.request)
        .then(response => {
          // Cache the fresh copy for offline use
          const clone = response.clone();
          caches.open(CACHE_VERSION).then(c => c.put(e.request, clone));
          return response;
        })
        .catch(() => {
          // Network genuinely failed → serve cached version
          return caches.match(e.request, { ignoreSearch: true })
            .then(cached => cached || caches.match('./app.html'));
        })
    );
    return;
  }

  // 4. Static assets (CSS, JS, images) → CACHE-FIRST with background refresh
  e.respondWith(
    caches.match(e.request, { ignoreSearch: true }).then(cached => {
      const networkFetch = fetch(e.request).then(response => {
        if (response && response.ok) {
          const clone = response.clone();
          caches.open(CACHE_VERSION).then(c => c.put(e.request, clone));
        }
        return response;
      }).catch(() => null);

      return cached || networkFetch;
    })
  );
});

// Listen for SKIP_WAITING message from the app
self.addEventListener('message', e => {
  if (e.data && e.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});
