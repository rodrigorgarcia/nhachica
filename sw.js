// Guarda o app no aparelho para abrir sem internet. Gerado por build.py.
var CACHE = "nhachica-37c637d71a";
var ARQUIVOS = ["./", "index.html", "manifest.webmanifest", "icon-192.png", "icon-512.png", "apple-touch-icon.png"];

self.addEventListener("install", function (e) {
  e.waitUntil(caches.open(CACHE).then(function (c) { return c.addAll(ARQUIVOS); }));
  self.skipWaiting();
});

self.addEventListener("activate", function (e) {
  e.waitUntil(caches.keys().then(function (chaves) {
    return Promise.all(chaves.filter(function (k) { return k !== CACHE; }).map(function (k) { return caches.delete(k); }));
  }));
  self.clients.claim();
});

// Responde do cache na hora e atualiza em segundo plano quando há sinal.
self.addEventListener("fetch", function (e) {
  if (e.request.method !== "GET") return;
  e.respondWith(caches.open(CACHE).then(function (c) {
    return c.match(e.request, { ignoreSearch: true }).then(function (guardado) {
      var rede = fetch(e.request).then(function (r) {
        if (r.ok || r.type === "opaque") c.put(e.request, r.clone());
        return r;
      }).catch(function () { return guardado; });
      return guardado || rede;
    });
  }));
});
