importScripts("https://cdn.onesignal.com/sdks/web/v16/OneSignalSDK.sw.js");

const CACHE = 'red-records-v10';
const PAGES = ['index.html', 'instrus.html', 'playlist.html', 'legal.html'];   // pages publiques (mises en cache)
const PRIVATE = ['admin.html', 'team.html'];                     // pages privees : jamais en cache, toujours reseau
const ASSETS = ['manifest.json', 'icons/icon-192.png', 'icons/icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll([...PAGES, ...ASSETS])));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);

  // Pages privees (dashboard, equipe) : toujours le reseau, jamais de cache.
  if (PRIVATE.some(p => url.pathname.endsWith(p))) {
    e.respondWith(fetch(e.request));
    return;
  }

  const isPage = PAGES.some(p => url.pathname.endsWith(p)) || url.pathname === '/';

  if (isPage) {
    // Réseau en priorité pour les pages → toujours la version à jour
    e.respondWith(
      fetch(e.request)
        .then(res => {
          const clone = res.clone();
          caches.open(CACHE).then(c => c.put(e.request, clone));
          return res;
        })
        .catch(() => caches.match(e.request))
    );
  } else {
    // Cache en priorité pour les assets (icônes, manifest, audio)
    e.respondWith(
      caches.match(e.request).then(cached => cached || fetch(e.request))
    );
  }
});
