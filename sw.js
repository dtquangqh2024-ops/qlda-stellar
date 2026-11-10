// Service worker tối thiểu để Android cài app như ứng dụng riêng. Không lưu đệm: luôn lấy bản mới nhất.
self.addEventListener('install', function () { self.skipWaiting(); });
self.addEventListener('activate', function (e) { e.waitUntil(self.clients.claim()); });
self.addEventListener('fetch', function (e) {
  var u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== self.location.origin) return; // dữ liệu Google để trình duyệt tự xử lý
  e.respondWith(fetch(e.request));
});
