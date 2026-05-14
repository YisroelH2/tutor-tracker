const CACHE_NAME = 'tutor-tracker-v1';
const urlsToCache = [
  '/',
  '/index.html',
  '/manifest.json',
  '/icon.png',
  'https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=Nunito:wght@400;600;700;800&display=swap',
  'https://cdnjs.cloudflare.com/ajax/libs/react/18.2.0/umd/react.production.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/react-dom/18.2.0/umd/react-dom.production.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/babel-standalone/7.23.2/babel.min.js',
];

// Install event: cache resources
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => {
        return cache.addAll(urlsToCache).catch((err) => {
          console.warn('Cache addAll error (some resources may not be available):', err);
        });
      })
  );
  self.skipWaiting();
});

// Activate event: clean up old caches
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((cacheName) => {
          if (cacheName !== CACHE_NAME) {
            return caches.delete(cacheName);
          }
        })
      );
    })
  );
  self.clients.claim();
});

// Fetch event: network first, fallback to cache
self.addEventListener('fetch', (event) => {
  // Skip non-GET requests
  if (event.request.method !== 'GET') return;

  // Network first for external APIs (Apps Script)
  if (event.request.url.includes('script.google.com')) {
    event.respondWith(
      fetch(event.request)
        .catch(() => {
          // If network fails, return a placeholder response
          return new Response(JSON.stringify({ error: 'offline' }), {
            status: 503,
            statusText: 'Service Unavailable',
            headers: new Headers({ 'Content-Type': 'application/json' }),
          });
        })
    );
    return;
  }

  // Cache first for everything else
  event.respondWith(
    caches.match(event.request)
      .then((response) => {
        return response || fetch(event.request)
          .then((response) => {
            // Cache successful responses
            if (!response || response.status !== 200 || response.type === 'error') {
              return response;
            }
            const responseToCache = response.clone();
            caches.open(CACHE_NAME)
              .then((cache) => {
                cache.put(event.request, responseToCache);
              });
            return response;
          })
          .catch(() => {
            // Return cached response or a fallback
            return caches.match(event.request)
              .then((cachedResponse) => {
                return cachedResponse || new Response('Offline - resource not available', {
                  status: 503,
                  statusText: 'Service Unavailable',
                });
              });
          });
      })
  );
});
