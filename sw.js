// ============================================================
//  Smart Attendance — Service Worker
//  Updated for AppStart Engine architecture.
//  Version is now managed via version.json (no config.js import).
// ============================================================

const CACHE_VERSION = 'attendance-v1.0.0';
const ASSETS = [
  './',
  './index.html',
  './css/app.css',
  './css/infographic.css',
  './js/api.js',
  './js/app.js',
  // AppStart Engine
  './appstart/config.js',
  './appstart/license.js',
  './appstart/keystore.js',
  './appstart/schema.js',
  './appstart/translator.js',
  './appstart/appstart.js',
  './appstart/appstart.css',
  // Assets
  './manifest.json',
  './version.json',
  './icons/icon-192.svg',
  './icons/icon-512.svg',
  './icons/rmdiper-logo.png',
];

// Install — cache core assets
self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE_VERSION)
      .then(cache => cache.addAll(ASSETS))
      .then(() => self.skipWaiting())
  );
});

// Activate — purge old caches & claim clients immediately
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys.filter(k => k !== CACHE_VERSION).map(k => caches.delete(k))
      )
    ).then(() => self.clients.claim())
  );
});

// Fetch — network-first for API & config sheet, cache-first for assets
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);

  // Dynamic external requests (Ad networks, APIs, CDNs, Google Sheets) → always network
  if (url.origin !== self.location.origin) {
    e.respondWith(fetch(e.request));
    return;
  }

  // version.json → always network (must bypass cache for update detection)
  if (url.pathname.endsWith('version.json')) {
    e.respondWith(
      fetch(e.request).catch(() => caches.match(e.request))
    );
    return;
  }

  // Assets → cache-first, fallback to network
  e.respondWith(
    caches.match(e.request).then(cached => {
      const fetchPromise = fetch(e.request).then(response => {
        if (response && response.status === 200) {
          const clone = response.clone();
          caches.open(CACHE_VERSION).then(cache => cache.put(e.request, clone));
        }
        return response;
      }).catch(() => cached);

      return cached || fetchPromise;
    })
  );
});

// Listen for update messages from the app
self.addEventListener('message', e => {
  if (e.data && e.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});
