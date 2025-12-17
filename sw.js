const CACHE_NAME = "karyamas-transkrip-v20";
const ASSETS = [
  "./",
  "./index.html",
  "./styles.css",
  "./manifest.webmanifest",
  "./assets/logo_karyamas.png",
  "./assets/logo_planters_academy.png",
  "./js/app.js"
];
self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE_NAME).then(c => c.addAll(ASSETS)).then(()=>self.skipWaiting()));
});
self.addEventListener("activate", (e) => {
  e.waitUntil(self.clients.claim());
});
self.addEventListener("fetch", (e) => {
  const req = e.request;
  e.respondWith(
    caches.match(req).then(res => res || fetch(req).then(netRes=>{
      const copy = netRes.clone();
      caches.open(CACHE_NAME).then(c=>c.put(req, copy)).catch(()=>{});
      return netRes;
    }).catch(()=>caches.match("./index.html")))
  );
});
