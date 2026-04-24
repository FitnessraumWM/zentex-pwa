const CACHE = "zentex-sprinkler-v1.3"; // 🔧 FIX 8 – Version hochgezählt

const ASSETS = [
  "./",
  "./index.html",
  "./offline.html",
  "./instandsetzung.html",
  "./checkliste.html",
  "./done.html",
  "./final.html",
  "./regie.html",
  "./Checkliste_rapport.html",
  "./regie_rapport.html",
  "./zentex-sp.js",
  "./manifest.webmanifest",
  "./assets/logo.jpg",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
  "./vendor/html2pdf.bundle.min.js"  // 🔧 FIX 8 – lokal gecachte html2pdf-Bibliothek
];

// ✅ ADDED – Robusteres Install: einzeln cachen, damit ein fehlendes Asset nicht alles blockiert
self.addEventListener("install", (e) => {
  e.waitUntil(
    caches.open(CACHE).then((c) => {
      return Promise.allSettled(
        ASSETS.map((url) =>
          c.add(url).catch((err) => {
            console.warn("SW: Cache fehlgeschlagen für", url, err);
          })
        )
      );
    }).then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))
    ).then(async () => {
      await self.clients.claim();
      const clients = await self.clients.matchAll({ includeUncontrolled: true });
      clients.forEach((client) => client.postMessage({ type: "sw-activated", cache: CACHE }));
    })
  );
});

self.addEventListener("fetch", (e) => {
  const req = e.request;
  if (req.method !== "GET") return;

  // ✅ ADDED – SharePoint REST API Requests NICHT cachen (POST wird schon oben gefiltert,
  // aber GET-Requests an /_api/ sollen auch nicht aus dem Cache kommen)
  const url = new URL(req.url);
  if (url.pathname.includes("/_api/")) return;

  const accept = req.headers.get("accept") || "";
  const isHtml = accept.includes("text/html");

  e.respondWith(
    caches.match(req).then((cached) => {
      if (cached) return cached;

      return fetch(req)
        .then((res) => {
          const okToCache =
            res &&
            res.status === 200 &&
            res.type === "basic" &&
            new URL(req.url).origin === self.location.origin;

          if (okToCache) {
            const copy = res.clone();
            caches.open(CACHE).then((c) => c.put(req, copy));
          }
          return res;
        })
        .catch(() => {
          if (isHtml) {
            return caches.match("./offline.html") || caches.match("./index.html");
          }
          return new Response("Offline", {
            status: 503,
            statusText: "Offline",
            headers: { "Content-Type": "text/plain; charset=utf-8" }
          });
        });
    })
  );
});
