// Service worker DCL : rend l'app installable (Android) et garde la page hors connexion.
const CACHE = 'dcl-v1';
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(['./', 'manifest.json', 'icons/icon-192.png', 'images/auto.jpg', 'images/moto.jpg']))); self.skipWaiting(); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k.startsWith('dcl-') && k !== CACHE).map(k => caches.delete(k))))); self.clients.claim(); });
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  const url = new URL(e.request.url);
  // Page : réseau d'abord (toujours la dernière version), cache si hors ligne
  if (e.request.mode === 'navigate') {
    e.respondWith(fetch(e.request).then(r => { const c = r.clone(); caches.open(CACHE).then(x => x.put(e.request, c)); return r; }).catch(() => caches.match(e.request).then(r => r || caches.match('./'))));
    return;
  }
  if (url.origin !== location.origin || url.pathname.endsWith('.mp4')) return;
  // Images/icônes : cache tout de suite, mise à jour en arrière-plan
  e.respondWith(caches.match(e.request).then(hit => {
    const net = fetch(e.request).then(res => { if (res.ok) { const c = res.clone(); caches.open(CACHE).then(x => x.put(e.request, c)); } return res; }).catch(() => hit);
    return hit || net;
  }));
});
