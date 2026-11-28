// Coquille de l'appli gardée sur le téléphone (jamais les données ni les appels au serveur)
const CACHE = 'ad-agence-v10';
self.addEventListener('install', e => { self.skipWaiting(); e.waitUntil(caches.open(CACHE).then(c => c.addAll(['./', 'index.html', 'logo.png', 'manifest.json']).catch(() => {}))); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(k => Promise.all(k.filter(x => x !== CACHE).map(x => caches.delete(x)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return;
  e.respondWith(caches.open(CACHE).then(c => c.match(e.request, { ignoreSearch: true }).then(r => {
    const net = fetch(e.request).then(n => { if (n && n.ok) c.put(e.request, n.clone()); return n; }).catch(() => r);
    return r || net;
  })));
});
