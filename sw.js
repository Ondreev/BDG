const CACHE = 'bdg-v4-119';
const PHOTO_CACHE = 'bdg-photos-v1';
const ASSETS = ['./', './index.html', './manifest.webmanifest', './icon.svg', './qrcode.min.js', './jsqr.min.js'];
const PHOTO_HOSTS = ['lh3.googleusercontent.com'];
// новые фото на Bunny отдаёт прокси Маркетплейса (/api/cdn-proxy) — кешируем так же,
// но только его, а не весь домен (там же живут API, которые кешировать нельзя)
const PHOTO_PROXY = { hosts: ['karatmarket.ru', 'delivery-food-puce.vercel.app'], path: '/api/cdn-proxy' };
const isPhotoRequest = url => PHOTO_HOSTS.includes(url.hostname) ||
  (PHOTO_PROXY.hosts.includes(url.hostname) && url.pathname === PHOTO_PROXY.path);

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
  if (isPhotoRequest(url)) {
    e.respondWith(
      caches.open(PHOTO_CACHE).then(cache =>
        cache.match(e.request).then(cached => {
          // <img> шлёт запрос "вслепую" (no-cors) — ответ непрозрачный, и сломанный ответ
          // (5xx прокси на холодном старте) не отличить от нормального, а в кэше он
          // застрял бы навсегда. Для нашего прокси повторяем запрос как обычный cors
          // (он отдаёт Access-Control-Allow-Origin: *) и кладём в кэш только успешный
          const isProxy = PHOTO_PROXY.hosts.includes(url.hostname);
          const network = fetch(isProxy ? new Request(e.request.url, { mode: 'cors' }) : e.request)
            .then(res => { if (res.ok || res.type === 'opaque') cache.put(e.request, res.clone()); return res; })
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
