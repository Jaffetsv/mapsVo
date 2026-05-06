const CACHE_NAME = 'dlsr-maps-v7-20260506';
const ASSETS = [
  './', './index.html', './acceso.html', './buscar.html', './contrato.html', './placa.html', './ubicacion.html', './campo.html', './favoritos.html', './estado.html', './convertidor.html',
  './styles.css?v=20260506-7', './app-core.js?v=20260506-7', './ui-helpers.js?v=20260506-7', './data-worker.js?v=20260506-7', './data-config.js?v=20260506-7',
  './logo.png', './logo_ui.png', './logo_fondo_blanco.png', './manifest.json'
];
self.addEventListener('install', event => { self.skipWaiting(); event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS).catch(() => null))); });
self.addEventListener('activate', event => { event.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  const isData = /Data%20|Data\s|\.xlsx$|data_.*\.json$/i.test(url.href);
  if (isData) {
    event.respondWith(fetch(req, { cache: 'no-store' }).catch(() => caches.match(req)));
    return;
  }
  event.respondWith(caches.match(req).then(cached => cached || fetch(req).then(res => { const copy = res.clone(); caches.open(CACHE_NAME).then(cache => cache.put(req, copy)).catch(() => null); return res; }).catch(() => cached)));
});
