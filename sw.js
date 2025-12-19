/* sw.js — Karyamas Transkrip Nilai (Chrome Mobile friendly) */

const CACHE_VERSION = "v51"; // naikkan jika ada perubahan aset
const CACHE_NAME = `karyamas-transkrip-${CACHE_VERSION}`;

// Pakai path absolut (lebih konsisten daripada "./")
const ASSETS = [
  "/",
  "/index.html",
  "/styles.css",
  "/manifest.webmanifest",
  "/assets/logo_karyamas.png",
  "/assets/logo_planters_academy.png",
  "/js/app.js",
];

// Helper: buat Response fallback index
async function fallbackIndex() {
  const cache = await caches.open(CACHE_NAME);
  const cached = await cache.match("/index.html");
  return cached || new Response("Offline", { status: 503, headers: { "Content-Type": "text/plain" } });
}

// INSTALL: precache aset inti
self.addEventListener("install", (event) => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE_NAME);

    // addAll bisa gagal jika 1 file gagal. Jadi: fetch satu2 agar lebih tahan banting.
    await Promise.all(
      ASSETS.map(async (path) => {
        try {
          const req = new Request(path, { cache: "reload" });
          const res = await fetch(req);
          if (res && res.ok) await cache.put(path, res.clone());
        } catch (e) {
          // jangan gagalkan install hanya karena 1 aset gagal
        }
      })
    );

    await self.skipWaiting();
  })());
});

// ACTIVATE: bersihkan cache lama + claim clients
self.addEventListener("activate", (event) => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.map((k) => (k !== CACHE_NAME ? caches.delete(k) : Promise.resolve())));
    await self.clients.claim();
  })());
});

self.addEventListener("fetch", (event) => {
  const req = event.request;

  // hanya handle GET
  if (req.method !== "GET") return;

  // Hindari request Range (lebih aman untuk cache)
  if (req.headers.has("range")) return;

  const url = new URL(req.url);

  // ✅ jangan intercept cross-origin (CDN, GAS, dll)
  if (url.origin !== self.location.origin) return;

  // Normalisasi pathname
  const path = url.pathname;

  // =========================
  // 1) NAVIGASI (SPA)
  // Network-first, fallback cache index
  // =========================
  if (req.mode === "navigate") {
    event.respondWith((async () => {
      try {
        // Network-first
        const fresh = await fetch(req);
        // Simpan index.html saja (bukan semua navigasi path)
        if (fresh && fresh.ok) {
          const cache = await caches.open(CACHE_NAME);
          cache.put("/index.html", fresh.clone()).catch(() => {});
        }
        return fresh;
      } catch (e) {
        // Offline fallback
        return await fallbackIndex();
      }
    })());
    return;
  }

  // =========================
  // 2) ASSET STATIC (JS/CSS/IMG/manifest)
  // Stale-while-revalidate
  // =========================
  const isStaticAsset =
    path.endsWith(".js") ||
    path.endsWith(".css") ||
    path.endsWith(".png") ||
    path.endsWith(".jpg") ||
    path.endsWith(".jpeg") ||
    path.endsWith(".webp") ||
    path.endsWith(".svg") ||
    path.endsWith(".ico") ||
    path.endsWith(".webmanifest");

  if (isStaticAsset || ASSETS.includes(path)) {
    event.respondWith((async () => {
      const cache = await caches.open(CACHE_NAME);
      const cached = await cache.match(path);

      const fetchPromise = fetch(req)
        .then((res) => {
          if (res && res.ok) cache.put(path, res.clone()).catch(() => {});
          return res;
        })
        .catch(() => null);

      // tampilkan cache dulu (kalau ada), sambil update di belakang
      return cached || (await fetchPromise) || new Response("Offline", { status: 503 });
    })());
    return;
  }

  // =========================
  // 3) DEFAULT: network-first, fallback cache
  // =========================
  event.respondWith((async () => {
    try {
      const res = await fetch(req);
      return res;
    } catch (e) {
      const cache = await caches.open(CACHE_NAME);
      const cached = await cache.match(path);
      return cached || new Response("Offline", { status: 503 });
    }
  })());
});
