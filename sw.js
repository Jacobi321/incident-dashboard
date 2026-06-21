var CACHE_NAME = 'osogbo-rcc-dashboard-v1';
var APP_SHELL = ['./', './index.html', './manifest.json', './icon.svg'];

self.addEventListener('install', function (event) {
  event.waitUntil(
    caches.open(CACHE_NAME).then(function (cache) { return cache.addAll(APP_SHELL); })
  );
  self.skipWaiting();
});

self.addEventListener('activate', function (event) {
  event.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.filter(function (k) { return k !== CACHE_NAME; }).map(function (k) { return caches.delete(k); }));
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', function (event) {
  var url = new URL(event.request.url);

  // Always go to the network for the live data API and any third-party
  // CDN assets - only the static app shell is cached for offline use.
  if (url.origin.indexOf('script.google') !== -1 ||
      url.origin.indexOf('googleusercontent') !== -1 ||
      url.origin.indexOf(self.location.origin) === -1) {
    return;
  }

  event.respondWith(
    caches.match(event.request).then(function (cached) {
      return cached || fetch(event.request);
    })
  );
});
