Claude




























Ssr sw · JS
/* ssr-sw.js : SSR温測 アプリ化用の裏方
 * 通信できるときは常に最新のファイルを取得し、通信できないときだけ保存分を使う。
 * (古い画面が残り続けないための作りです)
 */
var CACHE = 'ssr-app-v1';
 
self.addEventListener('install', function () {
  self.skipWaiting();
});
 
self.addEventListener('activate', function (e) {
  e.waitUntil(self.clients.claim());
});
 
self.addEventListener('fetch', function (e) {
  if (e.request.method !== 'GET') return;
  var sameOrigin = new URL(e.request.url).origin === self.location.origin;
  e.respondWith(
    fetch(e.request)
      .then(function (res) {
        if (sameOrigin && res && res.ok) {
          var copy = res.clone();
          caches.open(CACHE).then(function (c) {
            c.put(e.request, copy);
          });
        }
        return res;
      })
      .catch(function () {
        return caches.match(e.request);
      })
  );
});
 
Claude が応答を完了しました
