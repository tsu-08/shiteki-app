// Service Worker - 指摘写真生成アプリ
const CACHE_NAME = 'shitekishashin-v1';
const ASSETS = [
  './index.html',
  './manifest.json',
  './sw.js',
  'https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;600;700;900&display=swap',
  'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'
];

// インストール時にキャッシュ
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      // 外部リソースはキャッシュ失敗しても続行
      return cache.addAll(ASSETS).catch(() => {
        return cache.add('./指摘写真生成アプリv5.html');
      });
    })
  );
  self.skipWaiting();
});

// 古いキャッシュを削除
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// リクエスト時：キャッシュ優先、なければネットワーク
self.addEventListener('fetch', event => {
  // GASへのPOSTリクエストはキャッシュしない
  if (event.request.method === 'POST') return;

  event.respondWith(
    caches.match(event.request).then(cached => {
      if (cached) return cached;
      return fetch(event.request).then(response => {
        // 成功したレスポンスをキャッシュに追加
        if (response && response.status === 200 && response.type !== 'opaque') {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        }
        return response;
      }).catch(() => {
        // オフライン時はHTMLを返す
        if (event.request.destination === 'document') {
          return caches.match('./index.html');
        }
      });
    })
  );
});
