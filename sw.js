const CACHE_NAME = "acciones-app-v1";
const ASSETS = [
  "./canvas_acciones_final_pwa.html",
  "./canvas_acciones_movil.html",
  "./manifest.webmanifest",
  "./icon-192.svg",
  "./icon-512.svg"
];
self.addEventListener("install", event => {
  event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS)));
  self.skipWaiting();
});
self.addEventListener("activate", event => {
  event.waitUntil(self.clients.claim());
});
self.addEventListener("fetch", event => {
  event.respondWith(
    caches.match(event.request).then(cached => cached || fetch(event.request))
  );
});
