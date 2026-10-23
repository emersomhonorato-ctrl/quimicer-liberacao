// Service worker do app Liberação de Lotes Quimicer.
// Sempre busca a versão mais nova na internet; o cache só é usado se a conexão cair.
const CACHE = 'quimicer-v1';
const BASE = ['./', 'index.html', 'config.js', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png'];
self.addEventListener('install', (e) => { self.skipWaiting(); e.waitUntil(caches.open(CACHE).then((c) => c.addAll(BASE)).catch(() => {})); });
self.addEventListener('activate', (e) => { e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', (e) => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== self.location.origin) return; // Supabase e CDNs passam direto
  e.respondWith(fetch(e.request).then((r) => { const c = r.clone(); caches.open(CACHE).then((k) => k.put(e.request, c)).catch(() => {}); return r; })
    .catch(() => caches.match(e.request).then((m) => m || caches.match('index.html'))));
});
