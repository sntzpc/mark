const CACHE_NAME = "karyamas-transkrip-v26"; // naikkan versi agar update pasti jalan
const ASSETS = [
  "./",
  "./index.html",
  "./styles.css",
  "./manifest.webmanifest",
  "./assets/logo_karyamas.png",
  "./assets/logo_planters_academy.png",
  "./js/app.js"
];

// Install: pre-cache aset lokal
self.addEventListener("install", (e) => {
  e.waitUntil(
    caches.open(CACHE_NAME)
      .then(c => c.addAll(ASSETS))
      .then(() => self.skipWaiting())
  );
});

// Activate: hapus cache versi lama
self.addEventListener("activate", (e) => {
  e.waitUntil(
    Promise.all([
      caches.keys().then(keys =>
        Promise.all(keys.map(k => (k !== CACHE_NAME ? caches.delete(k) : Promise.resolve())))
      ),
      self.clients.claim()
    ])
  );
});

self.addEventListener("fetch", (e) => {
  const req = e.request;

  // hanya handle GET
  if (req.method !== "GET") return;

  const url = new URL(req.url);

  // ✅ PENTING: jangan intercept request cross-origin (CDN, GAS, dll)
  if (url.origin !== self.location.origin) {
    return; // biarkan browser fetch normal
  }

  // Navigasi halaman: network-first, fallback cache, terakhir index.html
  if (req.mode === "navigate") {
    e.respondWith(
      fetch(req)
        .then(res => {
          const copy = res.clone();
          caches.open(CACHE_NAME).then(c => c.put(req, copy)).catch(()=>{});
          return res;
        })
        .catch(() => caches.match(req).then(r => r || caches.match("./index.html")))
    );
    return;
  }

  // Asset lokal: cache-first lalu update cache dari network
  e.respondWith(
    caches.match(req).then(cached => {
      if (cached) return cached;

      return fetch(req).then(res => {
        const copy = res.clone();
        caches.open(CACHE_NAME).then(c => c.put(req, copy)).catch(()=>{});
        return res;
      });
    })
  );
});
