// BIB Loans — Service Worker v2
// Build, Invest, Borrow | Loans Made Simple

const CACHE_NAME = 'bib-loans-v3';
const STATIC_ASSETS = [
  './index.html',
  './manifest.json',
  './icon-192.jpg',
  './icon-512.jpg'
];

// ── Install: cache static shell ──────────────────────────────────
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      console.log('[BIB SW] Caching app shell');
      return cache.addAll(STATIC_ASSETS);
    })
  );
  self.skipWaiting();
});

// ── Activate: clean old caches ───────────────────────────────────
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys
          .filter((key) => key !== CACHE_NAME)
          .map((key) => {
            console.log('[BIB SW] Deleting old cache:', key);
            return caches.delete(key);
          })
      )
    )
  );
  self.clients.claim();
});

// ── Fetch: network-only for GAS API, cache-first for static ─────
self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);

  // Google Apps Script API calls — network only
  if (url.hostname.includes('script.google.com')) {
    event.respondWith(
      fetch(event.request).catch(() =>
        new Response(JSON.stringify({ error: 'Offline. Please reconnect.' }), {
          headers: { 'Content-Type': 'application/json' }
        })
      )
    );
    return;
  }

  // Google Fonts — network first, fallback to cache
  if (url.hostname.includes('fonts.googleapis.com') || url.hostname.includes('fonts.gstatic.com')) {
    event.respondWith(
      fetch(event.request)
        .then((response) => {
          const cloned = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, cloned));
          return response;
        })
        .catch(() => caches.match(event.request))
    );
    return;
  }

  // Static assets — cache first, then network
  event.respondWith(
    caches.match(event.request).then((cached) => {
      if (cached) return cached;
      return fetch(event.request).then((response) => {
        if (!response || response.status !== 200 || response.type !== 'basic') {
          return response;
        }
        const cloned = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, cloned));
        return response;
      });
    })
  );
});

// ── Background sync ──────────────────────────────────────────────
self.addEventListener('sync', (event) => {
  if (event.tag === 'bib-sync-loans') {
    event.waitUntil(syncPendingLoans());
  }
});

async function syncPendingLoans() {
  const clients = await self.clients.matchAll();
  clients.forEach((client) => client.postMessage({ type: 'SYNC_LOANS' }));
}
