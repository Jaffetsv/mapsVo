const CACHE_NAME = 'dlsr-maps-v20260506-3';
const STATIC_ASSETS = [
  './',
  './index.html',
  './acceso.html',
  './contrato.html',
  './placa.html',
  './ubicacion.html',
  './styles.css',
  './app-core.js',
  './manifest.json',
  './data-config.js',
  './logo.png',
  './logo_fondo_blanco.png'
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

function esDinamico(request) {
  const url = new URL(request.url);
  const path = url.pathname.toLowerCase();
  return request.mode === 'navigate' || path.endsWith('.html') || path.endsWith('.xlsx') || path.endsWith('data-config.js') || path.endsWith('app-core.js') || path.endsWith('styles.css');
}

async function networkFirst(request) {
  try {
    return await fetch(new Request(request, { cache: 'no-store' }));
  } catch (_) {
    const cached = await caches.match(request);
    if (cached) return cached;
    throw _;
  }
}

self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return;

  if (esDinamico(event.request)) {
    event.respondWith(networkFirst(event.request));
    return;
  }

  event.respondWith(
    caches.match(event.request).then(cached => {
      if (cached) return cached;
      return fetch(event.request).then(response => {
        if (response.ok && new URL(event.request.url).origin === self.location.origin) {
          const copy = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, copy));
        }
        return response;
      });
    })
  );
});
