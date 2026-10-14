// Karma Diagnose - service worker
// index-hash: 56301dc5f1
// Bump CACHE_VERSION every time index.html (or this file) changes.
// build.py does this automatically.
//
// Strategy: network-first. Always try the network first so updates take effect
// on the next reload; fall back to the cache only when offline.

const CACHE_VERSION = "v6";
const CACHE_NAME = "karma-diagnose-" + CACHE_VERSION;

const ASSETS = [
  "./",
  "./index.html",
  "./manifest.json",
  "./icon-180.png"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS))
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys
          .filter((key) => key.startsWith("karma-diagnose") && key !== CACHE_NAME)
          .map((key) => caches.delete(key))
      )
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") return;

  event.respondWith(
    fetch(event.request)
      .then((response) => {
        if (response && response.status === 200) {
          const copy = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
        }
        return response;
      })
      .catch(() => caches.match(event.request))
  );
});
