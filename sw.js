// 離線瀏覽：有網路時抓最新版並存起來，沒網路時用存好的版本
const CACHE = 'tokyo-trip-v1';
const PRECACHE = [
  './',
  './index.html',
  './sync-guide.html',
  './guide/step1.png',
  './guide/step2.png',
  './guide/step3.png',
  './guide/step4.png',
  './guide/step5.png',
  './guide/step6.png',
  './guide/step7.png'
];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(PRECACHE)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  // 只處理本網站的 GET；記帳同步（Google）不經過快取
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;
  e.respondWith(
    fetch(req)
      .then((res) => {
        if (res.ok) {
          const copy = res.clone();
          caches.open(CACHE).then((c) => c.put(req, copy));
        }
        return res;
      })
      .catch(() => caches.match(req, { ignoreSearch: true }).then((r) => r || (req.mode === 'navigate' ? caches.match('./index.html') : undefined)))
  );
});
