const CACHE = 'shamali-v3';
const ASSETS = [
  '/shamali-academy/',
  '/shamali-academy/index.html',
  'https://fonts.googleapis.com/css2?family=Cairo:wght@400;700;900&display=swap',
  'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.2/css/all.min.css'
];

self.addEventListener('install', function(e) {
  e.waitUntil(
    caches.open(CACHE).then(function(c) { return c.addAll(ASSETS); })
  );
  self.skipWaiting();
});

self.addEventListener('activate', function(e) {
  e.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(keys.filter(function(k){ return k !== CACHE; }).map(function(k){ return caches.delete(k); }));
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', function(e) {
  // للشبكة أولاً (بيانات Google Sheets)، ثم الكاش للملفات الثابتة
  if(e.request.url.includes('sheetdb.io') || e.request.url.includes('googleapis.com')) {
    e.respondWith(fetch(e.request).catch(function(){ return new Response('{}'); }));
    return;
  }
  e.respondWith(
    caches.match(e.request).then(function(cached) {
      return cached || fetch(e.request).then(function(res) {
        var clone = res.clone();
        caches.open(CACHE).then(function(c){ c.put(e.request, clone); });
        return res;
      });
    }).catch(function() {
      return caches.match('/shamali-academy/');
    })
  );
});
