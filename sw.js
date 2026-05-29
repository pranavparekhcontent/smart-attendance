// ============================================================
//  Smart Attendance — Service Worker  v1.0.2
//  FIXED: Cache version bumped → forces fresh install on all clients
//  FIXED: Fetch handler never returns undefined (prevents blank page)
//  FIXED: index.html cached for direct URL navigation
//  FIXED: Robust offline fallback
// ============================================================

const CACHE_VERSION = 'attendance-v1.0.4';  // ← BUMPED: forces old caches to clear
const ASSETS = [
  './',
  './index.html',   // ← ADDED: landing page must be cached too
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
  './version.json',
  './icons/icon-192.svg',
  './icons/icon-512.svg',
  './icons/icon-192.png',
  './icons/icon-512.png',
];

// Install — pre-cache all core assets
self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE_VERSION)
      .then(cache => cache.addAll(ASSETS))
      .then(() => self.skipWaiting())   // activate immediately
  );
});

// Activate — purge ALL old caches, claim all clients right away
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(
        keys.filter(k => k !== CACHE_VERSION).map(k => caches.delete(k))
      ))
      .then(() => self.clients.claim())
  );
});

// Fetch strategy:
//   - External requests (CDN, Google Sheets, APIs) → always network
//   - version.json → network-first (update detection must bypass cache)
//   - All other same-origin → cache-first with network update
//   - CRITICAL: always return a response, never undefined → prevents blank page
self.addEventListener('fetch', e => {
  // Fix for Chrome bug with DevTools / hard refresh
  if (e.request.cache === 'only-if-cached' && e.request.mode !== 'same-origin') {
    return;
  }

  const url = new URL(e.request.url);

  // 1. External requests → straight to network (no caching)
  if (url.origin !== self.location.origin) {
    e.respondWith(
      fetch(e.request).catch(() =>
        new Response('Network error', { status: 503, statusText: 'Offline' })
      )
    );
    return;
  }

  // 2. version.json → always network first, cache as fallback
  if (url.pathname.endsWith('version.json')) {
    e.respondWith(
      fetch(e.request)
        .then(r => {
          if (r.ok) {
            const clone = r.clone();
            caches.open(CACHE_VERSION).then(c => c.put(e.request, clone));
          }
          return r;
        })
        .catch(() => caches.match(e.request, { ignoreSearch: true }))
    );
    return;
  }

  // 3. Same-origin assets → cache-first, background network update
  e.respondWith(
    (async function() {
      try {
        const cached = await caches.match(e.request, { ignoreSearch: true });
        
        // Kick off a network fetch to keep the cache fresh
        const networkPromise = fetch(e.request).then(response => {
          if (response && response.ok) {
            const clone = response.clone();
            caches.open(CACHE_VERSION).then(cache => cache.put(e.request, clone));
          }
          return response;
        }).catch(() => null);  // network failed → null

        // Return cached immediately if available, otherwise wait for network
        if (cached) return cached;
        
        const networkResponse = await networkPromise;
        if (networkResponse) return networkResponse;
        
        // Both cache and network failed → show offline page for navigations
        if (e.request.mode === 'navigate' || (e.request.headers.get('accept') || '').includes('text/html')) {
          return new Response(
            `<!DOCTYPE html><html><head><meta charset="UTF-8">
            <meta name="viewport" content="width=device-width,initial-scale=1">
            <title>Offline — Smart Attendance</title>
            <style>body{background:#0D0F14;color:#E8EAF6;font-family:sans-serif;display:flex;
            align-items:center;justify-content:center;height:100vh;margin:0;text-align:center;}
            h2{color:#3B82F6;margin-bottom:8px;}p{color:#6B7280;font-size:14px;}
            button{margin-top:20px;background:#3B82F6;color:white;border:none;padding:12px 24px;
            border-radius:8px;cursor:pointer;font-size:15px;}</style></head><body>
            <div><h2>You're Offline</h2><p>Please check your connection and try again.</p>
            <button onclick="location.reload()">Try Again</button></div></body></html>`,
            { status: 503, headers: { 'Content-Type': 'text/html' } }
          );
        }
        return new Response('', { status: 404, statusText: 'Not Found' });
      } catch (err) {
        // ULTIMATE SAFETY NET: if ANYTHING crashes, try network, then show offline page
        console.error('SW fetch handler error:', err);
        try {
          return await fetch(e.request);
        } catch(e2) {
          return new Response(
            `<!DOCTYPE html><html><head><meta charset="UTF-8">
            <meta name="viewport" content="width=device-width,initial-scale=1">
            <title>Offline — Smart Attendance</title>
            <style>body{background:#0D0F14;color:#E8EAF6;font-family:sans-serif;display:flex;
            align-items:center;justify-content:center;height:100vh;margin:0;text-align:center;}
            h2{color:#3B82F6;margin-bottom:8px;}p{color:#6B7280;font-size:14px;}
            button{margin-top:20px;background:#3B82F6;color:white;border:none;padding:12px 24px;
            border-radius:8px;cursor:pointer;font-size:15px;}</style></head><body>
            <div><h2>You're Offline</h2><p>Please check your connection and try again.</p>
            <button onclick="location.reload()">Try Again</button></div></body></html>`,
            { status: 503, headers: { 'Content-Type': 'text/html' } }
          );
        }
      }
    })()
  );
});

// Listen for SKIP_WAITING message from the app
self.addEventListener('message', e => {
  if (e.data && e.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});
