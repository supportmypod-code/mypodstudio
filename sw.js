/* MyPod service worker. Bump VERSION on every release so customers get the new files. */
const VERSION = "mypod-v8";
const CORE = ["./", "index.html", "manifest.webmanifest", "logo.png", "icons/icon-192.png", "icons/icon-512.png"];
self.addEventListener("install", e => e.waitUntil(caches.open(VERSION).then(c => c.addAll(CORE)).then(() => self.skipWaiting())));
self.addEventListener("activate", e => e.waitUntil(caches.keys().then(k => Promise.all(k.filter(x => x !== VERSION).map(x => caches.delete(x)))).then(() => self.clients.claim())));
self.addEventListener("fetch", e => {
  const r = e.request;
  if (r.method !== "GET" || new URL(r.url).origin !== location.origin) return;   // bookings (Apps Script) and remote images always go to the network
  if (r.mode === "navigate") {                                                    // pages: network first, cached copy when offline
    e.respondWith(fetch(r).then(res => { if (res.ok) { const c = res.clone(); caches.open(VERSION).then(x => x.put("index.html", c)); } return res; }).catch(() => caches.match("index.html")));
    return;
  }
  e.respondWith(caches.match(r).then(hit => {                                      // files: cached copy now, refreshed in the background
    const net = fetch(r).then(res => { if (res.ok) { const c = res.clone(); caches.open(VERSION).then(x => x.put(r, c)); } return res; }).catch(() => hit);
    return hit || net;
  }));
});
