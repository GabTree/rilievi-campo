// Service worker: conserva l'app sul telefono per l'uso offline.
const CACHE = "rilievi-v0.5.9b";
const TILES = "rilievi-tiles";
let trimming = false;
async function trimTiles(c) {   // tiene al massimo ~1500 riquadri (circa 30 MB)
  if (trimming) return; trimming = true;
  const keys = await c.keys();
  for (let i = 0; i < keys.length - 1500; i++) await c.delete(keys[i]);
  trimming = false;
}
const SHELL = ["./", "index.html", "app.js", "config.js", "manifest.webmanifest", "specie.json",
  "vendor/msal-browser.min.js", "vendor/piexif.js", "vendor/jszip.min.js", "vendor/leaflet.js", "vendor/leaflet.css", "arete.js", "import.js", "install.js", "viewer.js", "tools.js", "output.js", "vendor/docx.min.js", "vendor/qrcode.js", "brand/logo_orizzontale.png", "arete_lists.json", "icons/icon-192.png", "icons/icon-512.png", "icons/apple-touch-icon.png", "icons/icon-maskable-512.png"];
self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== CACHE && k !== TILES).map((k) => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener("fetch", (e) => {
  const u = new URL(e.request.url);
  // riquadri della mappa: dalla memoria se già visti (anche offline), altrimenti dalla rete
  if (e.request.method === "GET" && u.hostname === "tile.openstreetmap.org") {
    e.respondWith(caches.open(TILES).then((c) => c.match(e.request).then((hit) => hit || fetch(e.request).then((r) => {
      if (r.ok) { c.put(e.request, r.clone()); trimTiles(c); }
      return r;
    }).catch(() => new Response("", { status: 504 })))));
    return;
  }
  if (e.request.method !== "GET" || u.origin !== location.origin) return;   // Graph, Google e login: sempre in rete
  // rete prima (per avere gli aggiornamenti), cache se offline
  // "no-cache": chiede sempre al server se il file è cambiato, così non si mischiano versioni diverse
  const req = e.request.mode === "navigate" ? new Request(e.request.url, { cache: "no-cache" }) : new Request(e.request, { cache: "no-cache" });
  e.respondWith(fetch(req).then((r) => {
    const copy = r.clone();
    caches.open(CACHE).then((c) => c.put(e.request, copy));
    return r;
  }).catch(() => caches.match(e.request, { ignoreSearch: true }).then((r) => r || caches.match("index.html"))));
});
