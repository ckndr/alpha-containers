// Tubex Service Worker v4
// Network-first: always fetches fresh data when online
// Falls back to cache only when offline

const CACHE_NAME = 'tubex-202609021349';
const ASSETS = [
  './',
  './index.html',
  './Tubex.html',
  './manifest.json',
  './icon-192-any.png',
  './icon-512-any.png',
  './icon-192-maskable.png',
  './icon-512-maskable.png',
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
  // Only handle GET requests with http/https schemes (Rule R3-05)
  if (event.request.method !== 'GET' || !event.request.url.startsWith('http')) return;

  event.respondWith(
    fetch(event.request)
      .then(response => {
        // Only cache successful 200 responses to prevent caching 404/500 errors (Rule R3-04)
        if (response && response.status === 200) {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        }
        return response;
      })
      .catch(() => {
        // Network failed — serve from cache
        return caches.match(event.request).then(cached => {
          if (cached) return cached;
          // Fallback to cached Tubex.html for HTML navigation requests
          if (event.request.mode === 'navigate' || event.request.headers.get('accept')?.includes('text/html')) {
            return caches.match('./Tubex.html');
          }
          // Nothing cached — return a simple offline message
          return new Response(
            '<h2 style="font-family:sans-serif;padding:20px">Offline — open when connected to see latest data.</h2>',
            { headers: { 'Content-Type': 'text/html' } }
          );
        });
      })
  );
});
