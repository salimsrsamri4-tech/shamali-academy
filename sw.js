const CACHE = 'shamali-v6';
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
  // الشبكة أولاً دائماً لأي بيانات خارجية
  if(
    e.request.url.includes('sheetdb.io') ||
    e.request.url.includes('googleapis.com') ||
    e.request.url.includes('script.google.com') ||
    e.request.url.includes('googleusercontent.com')
  ) {
    e.respondWith(fetch(e.request).catch(function(){ return new Response('[]'); }));
    return;
  }
  // الشبكة أولاً لـ index.html لضمان التحديث دائماً
  if(e.request.url.includes('index.html') || e.request.mode === 'navigate') {
    e.respondWith(
      fetch(e.request).then(function(res) {
        var clone = res.clone();
        caches.open(CACHE).then(function(c){ c.put(e.request, clone); });
        return res;
      }).catch(function() {
        return caches.match(e.request);
      })
    );
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
