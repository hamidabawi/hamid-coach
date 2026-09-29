/* Hamid Coach service worker.
   - App pages (HTML): network first, cached copy when offline.
   - Versioned assets (app.js?v=N, vendor libs, icons): cache first.
   - The Google Sheets API / anything cross-origin: never touched.
   To ship an update: change the ?v= number in index.html AND in SHELL below, and bump VERSION. */
const VERSION = 'coach-v9';
const SHELL = ['./', './app.js?v=9', './manifest.webmanifest', './icons/icon-192.png', './icons/icon-512.png', './icons/apple-touch-icon.png', './vendor/chart-4.5.1.umd.min.js'];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== VERSION).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request, url = new URL(req.url);
  if (req.method !== 'GET' || url.origin !== location.origin) return;
  if (req.mode === 'navigate') {
    const isApp = /\/(index\.html)?$/.test(url.pathname);
    e.respondWith(fetch(req).then(res => { if (res.ok && isApp) { const c = res.clone(); caches.open(VERSION).then(x => x.put('./', c)); } return res; }).catch(() => caches.match('./')));
    return;
  }
  e.respondWith(caches.match(req).then(hit => hit || fetch(req).then(res => { if (res.ok) { const c = res.clone(); caches.open(VERSION).then(x => x.put(req, c)); } return res; })));
});
