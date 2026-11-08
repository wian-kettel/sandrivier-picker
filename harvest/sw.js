// Lets the Harvest Tracker open with no signal (the FORT 1 in the orchard).
// App files: network first, phone's copy when offline. Firebase and Excel libraries: kept once downloaded.
// Data itself is handled by Firebase's own offline storage, not here.
const CACHE = 'harvest-v8';
const FILES = ['./', 'index.html', 'config.js', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png'];
const LIBS = /^https:\/\/(www\.gstatic\.com\/firebasejs\/|cdn\.jsdelivr\.net\/npm\/xlsx@)/;
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES))); self.skipWaiting(); });
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k.indexOf('harvest-') === 0 && k !== CACHE).map(k => caches.delete(k)))));
  self.clients.claim();
});
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  const url = new URL(e.request.url);
  if (LIBS.test(e.request.url)) {
    e.respondWith(caches.match(e.request).then(hit => hit || fetch(e.request).then(r => {
      const copy = r.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); return r; })));
    return;
  }
  if (url.origin !== location.origin) return;            // never touch Firebase data or login calls
  const net = fetch(e.request).then(r => {
    if (r.ok) { const copy = r.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); }
    return r;
  });
  const cached = () => caches.match(e.request, { ignoreSearch: true }).then(r => r || caches.match('index.html'));
  // weak signal in the orchard: if the network takes more than 4 s, open from the phone's copy
  const slow = new Promise(res => setTimeout(res, 4000)).then(cached).then(r => r || net);
  e.respondWith(Promise.race([net.catch(cached), slow]));
});
