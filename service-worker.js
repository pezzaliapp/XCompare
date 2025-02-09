const CACHE_NAME = "xcompare-cache-v2"; // Cambia la versione per forzare aggiornamento
const urlsToCache = [
  "/",
  "/index.html",
  "/styles.css",
  "/app.js",
  "/manifest.json",
  "/XTM-192.png",
  "/XTM-512.png"
];

// Install Service Worker
self.addEventListener("install", (event) => {
  console.log("[Service Worker] Installing...");
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      console.log("[Service Worker] Caching files...");
      return cache.addAll(urlsToCache);
    })
  );
});

// Fetch Requests
self.addEventListener("fetch", (event) => {
  console.log("[Service Worker] Fetching:", event.request.url);
  event.respondWith(
    caches.match(event.request).then((response) => {
      return response || fetch(event.request);
    })
  );
});

// Activate Service Worker
self.addEventListener("activate", (event) => {
  console.log("[Service Worker] Activating new Service Worker...");
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((cache) => {
          if (cache !== CACHE_NAME) {
            console.log("[Service Worker] Deleting old cache:", cache);
            return caches.delete(cache);
          }
        })
      );
    })
  );
});