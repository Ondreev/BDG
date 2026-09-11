const CACHE = 'bdg-v4-109';
const PHOTO_CACHE = 'bdg-photos-v1';
const ASSETS = ['./', './index.html', './manifest.webmanifest', './icon.svg', './qrcode.min.js', './jsqr.min.js'];
const PHOTO_HOSTS = ['lh3.googleusercontent.com'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      // PHOTO_CACHE не привязан к версии приложения — фото не должны стираться при каждом
      // обновлении sw.js, иначе кэш фото товаров был бы бесполезен
      .then(keys => Promise.all(keys.filter(k => k !== CACHE && k !== PHOTO_CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  const url = new URL(e.request.url);

  // фото товаров с Google Диска: снимок один раз загруженный меняется редко —
  // отдаём мгновенно из кэша, если он есть, и параллельно тихо обновляем его на
  // будущее в фоне (stale-while-revalidate), а не ждём сеть на каждый повторный показ
  if (PHOTO_HOSTS.includes(url.hostname)) {
    e.respondWith(
      caches.open(PHOTO_CACHE).then(cache =>
        cache.match(e.request).then(cached => {
          const network = fetch(e.request)
            .then(res => { cache.put(e.request, res.clone()); return res; })
            .catch(() => null);
          return cached || network;
        })
      )
    );
    return;
  }

  if (url.origin !== location.origin) return;
  // network-first so updates arrive quickly, cache fallback for offline
  e.respondWith(
    fetch(e.request)
      .then(res => {
        const copy = res.clone();
        caches.open(CACHE).then(c => c.put(e.request, copy));
        return res;
      })
      .catch(() => caches.match(e.request).then(m => m || caches.match('./index.html')))
  );
});
