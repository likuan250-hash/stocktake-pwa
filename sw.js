const CACHE = 'stocktake-pwa-v33';
const ASSETS = [
  './', './index.html',
  './css/styles.css',
  './js/db-core.js', './js/db.js', './js/import.js', './js/export.js', './js/kingdee-sheets.js', './js/app.js',
  './vendor/sql-wasm.js', './vendor/sql-wasm.wasm', './vendor/xlsx.full.min.js',
  './manifest.webmanifest',
  './assets/icon-192.png', './assets/icon-512.png', './assets/apple-touch-icon.png'
];
// 数据文件走 network-first：确保老用户每次都能拿到最新金蝶同步库，离线时才回退缓存
const NETWORK_FIRST = ['kingdee-sheets.js'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  const path = new URL(e.request.url).pathname.split('/').pop();

  if (NETWORK_FIRST.indexOf(path) >= 0) {
    e.respondWith(
      fetch(e.request).then(resp => {
        const cp = resp.clone();
        caches.open(CACHE).then(c => c.put(e.request, cp)).catch(() => {});
        return resp;
      }).catch(() => caches.match(e.request).then(r => r || caches.match('./')))
    );
    return;
  }

  // 其余静态资源：cache-first
  e.respondWith(
    caches.match(e.request).then(r =>
      r || fetch(e.request).then(resp => {
        const cp = resp.clone();
        caches.open(CACHE).then(c => c.put(e.request, cp)).catch(() => {});
        return resp;
      }).catch(() => r)
    )
  );
});
