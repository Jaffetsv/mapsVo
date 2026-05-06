const CACHE_NAME = 'dlsr-maps-v20260506-6';
const STATIC_ASSETS = [
  './logo_ui.png',
  './logo.png',
  './logo_fondo_blanco.png',
  './manifest.json'
];

self.addEventListener('install', event => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(STATIC_ASSETS).catch(() => undefined))
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

function shouldNetworkFirst(request) {
  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return false;
  const path = url.pathname.toLowerCase();
  return request.mode === 'navigate' ||
    path.endsWith('.html') ||
    path.endsWith('.xlsx') ||
    path.endsWith('.js') ||
    path.endsWith('.css') ||
    path.endsWith('.json');
}

async function networkFirst(request) {
  try {
    const fresh = await fetch(new Request(request, { cache: 'no-store' }));
    return fresh;
  } catch (err) {
    const cached = await caches.match(request);
    if (cached) return cached;
    throw err;
  }
}

self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return;

  if (shouldNetworkFirst(event.request)) {
    event.respondWith(networkFirst(event.request));
    return;
  }

  event.respondWith(
    caches.match(event.request).then(cached => {
      if (cached) return cached;
      return fetch(event.request).then(response => {
        const url = new URL(event.request.url);
        if (response.ok && url.origin === self.location.origin) {
          const copy = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, copy));
        }
        return response;
      });
    })
  );
});
