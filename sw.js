const CACHE_NAME = "daily-ledger-cache-v8";
const APP_SHELL = [
  "./",
  "./index.html",
  "./style.css",
  "./app.js",
  "./firebase-config.js",
  "./manifest.json",
  "./icons/icon-192.png",
  "./icons/icon-512.png"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL))
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  const req = event.request;
  // Network-first for the SheetJS + Firebase SDK CDN scripts so the app always
  // gets the latest library build when online, falling back to cache offline.
  // (Note: actual database reads/writes go over Firebase's own WebSocket
  // connection, not plain fetches, so this only covers loading the SDK itself —
  // you still need connectivity to sign in and sync data.)
  if (req.url.includes("cdnjs.cloudflare.com") || req.url.includes("gstatic.com")) {
    event.respondWith(
      fetch(req)
        .then((res) => {
          const clone = res.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(req, clone));
          return res;
        })
        .catch(() => caches.match(req))
    );
    return;
  }
  // Cache-first for the app shell.
  event.respondWith(
    caches.match(req).then((cached) => cached || fetch(req))
  );
});
