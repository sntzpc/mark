/* sw.js — Karyamas Transkrip Nilai (Offline-first, Chrome Mobile friendly)
   - Aman untuk deploy di root maupun subfolder
   - Tidak intercept cross-origin (GAS/CDN)
   - SPA navigation fallback ke index.html
   - Static assets: stale-while-revalidate
*/

const CACHE_VERSION = "v30"; // naikkan jika ada perubahan aset
const CACHE_NAME = `karyamas-transkrip-${CACHE_VERSION}`;

// Base path tempat sw.js berada (mendukung deploy subfolder)
const BASE = self.location.pathname.replace(/\/sw\.js$/, "");

// Daftar aset inti untuk precache
const ASSETS = [
  `${BASE}/`,
  `${BASE}/index.html`,
  `${BASE}/styles.css`,
  `${BASE}/manifest.webmanifest`,
  `${BASE}/assets/logo_karyamas.png`,
  `${BASE}/assets/logo_planters_academy.png`,
  `${BASE}/js/app.js`,
];

// Fallback index untuk offline navigasi
async function fallbackIndex() {
  const cache = await caches.open(CACHE_NAME);
  const cached = await cache.match(`${BASE}/index.html`);
  return (
    cached ||
    new Response("Offline", {
      status: 503,
      headers: { "Content-Type": "text/plain" },
    })
  );
}

// INSTALL: precache aset inti (tahan banting)
self.addEventListener("install", (event) => {
  event.waitUntil(
    (async () => {
      const cache = await caches.open(CACHE_NAME);

      await Promise.all(
        ASSETS.map(async (path) => {
          try {
            const req = new Request(path, { cache: "reload" });
            const res = await fetch(req);
            if (res && res.ok) await cache.put(req, res.clone());
          } catch (e) {
            // jangan gagalkan install hanya karena 1 aset gagal
          }
        })
      );

      await self.skipWaiting();
    })()
  );
});

// ACTIVATE: bersihkan cache lama + claim clients
self.addEventListener("activate", (event) => {
  event.waitUntil(
    (async () => {
      const keys = await caches.keys();
      await Promise.all(keys.map((k) => (k !== CACHE_NAME ? caches.delete(k) : Promise.resolve())));
      await self.clients.claim();
    })()
  );
});

// (Opsional) dukung update instan via postMessage("SKIP_WAITING")
self.addEventListener("message", (event) => {
  if (event.data === "SKIP_WAITING") self.skipWaiting();
});

self.addEventListener("fetch", (event) => {
  const req = event.request;

  // hanya handle GET
  if (req.method !== "GET") return;

  // Hindari request Range (lebih aman untuk cache)
  if (req.headers.has("range")) return;

  const url = new URL(req.url);

  // Jangan intercept cross-origin (CDN, GAS, dll)
  if (url.origin !== self.location.origin) return;

  const path = url.pathname;

  // Jangan cache sourcemap
  if (path.endsWith(".map")) return;

  // =========================
  // 1) NAVIGASI (SPA)
  // Network-first, fallback cache index
  // =========================
  if (req.mode === "navigate") {
    event.respondWith(
      (async () => {
        try {
          const fresh = await fetch(req);
          // Simpan index.html saja
          if (fresh && fresh.ok) {
            const cache = await caches.open(CACHE_NAME);
            cache.put(new Request(`${BASE}/index.html`), fresh.clone()).catch(() => {});
          }
          return fresh;
        } catch (e) {
          return await fallbackIndex();
        }
      })()
    );
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

  // asset penting yang dipre-cache (cek berbasis path)
  const isInAssetList = ASSETS.includes(path) || ASSETS.includes(`${BASE}${path}`);

  if (isStaticAsset || isInAssetList) {
    event.respondWith(
      (async () => {
        const cache = await caches.open(CACHE_NAME);
        const cached = await cache.match(req);

        const fetchPromise = fetch(req)
          .then((res) => {
            if (res && res.ok) cache.put(req, res.clone()).catch(() => {});
            return res;
          })
          .catch(() => null);

        return cached || (await fetchPromise) || new Response("Offline", { status: 503 });
      })()
    );
    return;
  }

  // =========================
  // 3) DEFAULT: network-first, fallback cache
  // =========================
  event.respondWith(
    (async () => {
      try {
        return await fetch(req);
      } catch (e) {
        const cache = await caches.open(CACHE_NAME);
        const cached = await cache.match(req);
        return cached || new Response("Offline", { status: 503 });
      }
    })()
  );
});
