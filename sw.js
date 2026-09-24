/* Service worker de SIGA.
   Estrategia: primero la red (siempre trae lo último) y, si no hay internet,
   responde desde la caché. Así la app funciona también sin conexión. */
var CACHE = "siga-v1";
var ARCHIVOS = [
  "./",
  "./index.html",
  "./style.css",
  "./app.js",
  "./jszip.min.js",
  "./jspdf.umd.min.js",
  "./sena_logo.png",
  "./Siga.xlsx",
  "./Siga_planilla.xlsx",
  "./manifest.json",
  "./icon-192.png",
  "./icon-512.png"
];

self.addEventListener("install", function (e) {
  e.waitUntil(
    caches.open(CACHE).then(function (c) {
      return Promise.all(ARCHIVOS.map(function (u) {
        return c.add(u).catch(function () { /* si falta un archivo, no rompe la instalación */ });
      }));
    }).then(function () { return self.skipWaiting(); })
  );
});

self.addEventListener("activate", function (e) {
  e.waitUntil(
    caches.keys().then(function (claves) {
      return Promise.all(claves.map(function (k) {
        if (k !== CACHE) return caches.delete(k);
      }));
    }).then(function () { return self.clients.claim(); })
  );
});

self.addEventListener("fetch", function (e) {
  var req = e.request;
  if (req.method !== "GET") return;                  // el POST de conversión no se toca

  var url;
  try { url = new URL(req.url); } catch (err) { return; }
  if (url.origin !== self.location.origin) return;   // no interceptar servidores externos
  if (url.pathname.indexOf("/api/") === 0) return;   // sonda del conversor: siempre a la red

  e.respondWith(
    fetch(req).then(function (r) {
      if (r && r.ok) {
        var copia = r.clone();
        caches.open(CACHE).then(function (c) { c.put(req, copia); });
      }
      return r;
    }).catch(function () {
      return caches.match(req).then(function (m) {
        return m || caches.match("./index.html");
      });
    })
  );
});
