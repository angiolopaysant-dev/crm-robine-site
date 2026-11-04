/* Service worker du CRM : l’app s’ouvre hors connexion avec les dernières données (toujours chiffrées). */
const SHELL = "crm-shell-v1", DATA = "crm-data-v1", LIBS = "crm-libs-v1";
const FILES = ["./", "index.html", "style.css", "app.js", "crypto.js", "manifest.webmanifest", "icons/icon-192.png", "icons/icon-512.png", "icons/apple-touch-icon.png"];
self.addEventListener("install", e => { e.waitUntil(caches.open(SHELL).then(c => c.addAll(FILES)).then(() => self.skipWaiting())); });
self.addEventListener("activate", e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => ![SHELL, DATA, LIBS].includes(k)).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
const strip = req => { const u = new URL(req.url); u.search = ""; return u.toString(); };
self.addEventListener("fetch", e => {
  const req = e.request; if (req.method !== "GET") return;
  const u = new URL(req.url);
  if (u.origin === location.origin) {
    const p = u.pathname;
    if (/\/data\/f\/[0-9a-f]{32}$/.test(p)) { // fichiers chiffrés : contenu immuable
      e.respondWith(caches.open(DATA).then(async c => (await c.match(req)) || fetch(req).then(r => { if (r.ok) c.put(req, r.clone()); return r; })));
      return;
    }
    if (/\/(data\/data\.json|users\.json)$/.test(p)) { // données : réseau d’abord, copie chiffrée hors ligne
      e.respondWith(fetch(req).then(r => { if (r.ok) { const cl = r.clone(); caches.open(DATA).then(c => c.put(strip(req), cl)); } return r; })
        .catch(() => caches.open(DATA).then(c => c.match(strip(req))).then(r => r || Response.error())));
      return;
    }
    e.respondWith(caches.open(SHELL).then(async c => { // application : cache puis mise à jour en arrière-plan
      const hit = await c.match(req, { ignoreSearch: true });
      const net = fetch(req).then(r => { if (r.ok) c.put(req, r.clone()); return r; }).catch(() => hit || (req.mode === "navigate" ? c.match("index.html") : Response.error()));
      return hit || net;
    }));
    return;
  }
  if (/^(cdnjs\.cloudflare\.com|cdn\.jsdelivr\.net|fonts\.(googleapis|gstatic)\.com)$/.test(u.hostname)) {
    e.respondWith(caches.open(LIBS).then(async c => (await c.match(req)) || fetch(req).then(r => { if (r.ok || r.type === "opaque") c.put(req, r.clone()); return r; })));
  }
});
