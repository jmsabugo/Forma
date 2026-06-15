// Service worker: cachea el armazón de la app para funcionar sin cobertura.
// Las llamadas a Dropbox nunca se cachean.
const CACHE = 'forma-v23';
// Archivos propios: deben cachearse para que la instalación termine.
const LOCAL = [
  './',
  'index.html',
  'styles.css',
  'app.js',
  'manifest.webmanifest',
  'icon-192.png',
  'icon-512.png',
  'apple-touch-icon.png',
];
// CDN (SheetJS, Chart.js): se cachean en segundo plano; si fallan no bloquean.
const CDN = [
  'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js',
  'https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js',
];

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE)
      .then((c) => c.addAll(LOCAL))          // local: bloquea (rápido)
      .then(() => self.skipWaiting())
  );
  // CDN: best-effort, sin bloquear la instalación
  caches.open(CACHE).then((c) => CDN.forEach((u) => c.add(u).catch(() => {})));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  const url = e.request.url;
  if (url.includes('dropboxapi.com')) return; // siempre red
  e.respondWith(
    caches.match(e.request).then((hit) => hit ||
      fetch(e.request).then((r) => {
        if (e.request.method === 'GET' && r.ok) {
          const copia = r.clone();
          caches.open(CACHE).then((c) => c.put(e.request, copia));
        }
        return r;
      })
    )
  );
});
