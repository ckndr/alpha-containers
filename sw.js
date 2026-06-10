// Tubex Service Worker v3
// Network-first: always fetches fresh data when online
// Falls back to cache only when offline

const CACHE_NAME = 'tubex-202606101449';
const ASSETS = [
  './Tubex.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
];

// Install: pre-cache the app shell
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS))
  );
  self.skipWaiting();
});

// Activate: delete old caches
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))
      )
    )
  );
  self.clients.claim();
});

// Fetch: network-first, cache fallback
self.addEventListener('fetch', event => {
  // Only handle GET requests for our own origin
  if (event.request.method !== 'GET') return;

  event.respondWith(
    fetch(event.request)
      .then(response => {
        // Clone before consuming
        const clone = response.clone();
        caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        return response;
      })
      .catch(() => {
        // Network failed — serve from cache
        return caches.match(event.request).then(cached => {
          if (cached) return cached;
          // Nothing cached — return a simple offline message
          return new Response(
            '<h2 style="font-family:sans-serif;padding:20px">Offline — open when connected to see latest data.</h2>',
            { headers: { 'Content-Type': 'text/html' } }
          );
        });
      })
  );
});
