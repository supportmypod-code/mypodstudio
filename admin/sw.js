/* MyPod Admin service worker: caches the app shell so the app opens offline.
   Booking data and Google sign-in are never cached here. */
const VERSION = 'mypod-admin-v4';
const SHELL = ['./', './index.html', './manifest.webmanifest', './icon-192.png', './icon-512.png', './apple-touch-icon.png', './favicon-32.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== VERSION).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  // Never touch the booking API or Google sign-in
  if (/(^|\.)script\.google\.com$|(^|\.)script\.googleusercontent\.com$|(^|\.)accounts\.google\.com$/.test(url.hostname)) return;

  // Fonts: cache after first load
  if (url.hostname === 'fonts.googleapis.com' || url.hostname === 'fonts.gstatic.com') {
    e.respondWith(caches.open(VERSION).then(async c => {
      const hit = await c.match(req);
      if (hit) return hit;
      try { const res = await fetch(req); c.put(req, res.clone()); return res; } catch (err) { return hit || Response.error(); }
    }));
    return;
  }

  if (url.origin !== location.origin) return;

  // Page: network first so updates arrive, cached copy when offline
  if (req.mode === 'navigate') {
    e.respondWith(
      fetch(req).then(res => { const copy = res.clone(); caches.open(VERSION).then(c => c.put('./index.html', copy)); return res; })
        .catch(() => caches.match('./index.html'))
    );
    return;
  }

  // Other same-origin files: cached first, refreshed in the background
  e.respondWith(caches.match(req).then(hit => {
    const net = fetch(req).then(res => { caches.open(VERSION).then(c => c.put(req, res.clone())); return res; }).catch(() => hit);
    return hit || net;
  }));
});



/* Tapping a notification opens (or focuses) the admin app. */
self.addEventListener("notificationclick", e => {
  e.notification.close();
  e.waitUntil((async () => {
    const all = await clients.matchAll({ type: "window", includeUncontrolled: true });
    for (const c of all) { if ("focus" in c) { await c.focus(); return; } }
    if (clients.openWindow) await clients.openWindow("./");
  })());
});

/* Background check for new orders (Chrome / Edge on Android, installed app only). */
function mpIdb(mode, fn) {
  return new Promise((res, rej) => {
    const r = indexedDB.open("mpadmin", 1);
    r.onupgradeneeded = () => r.result.createObjectStore("kv");
    r.onerror = () => rej(r.error);
    r.onsuccess = () => {
      const tx = r.result.transaction("kv", mode), st = tx.objectStore("kv"), q = fn(st);
      tx.oncomplete = () => res(q && q.result);
      tx.onerror = () => rej(tx.error);
    };
  });
}
async function mpCheckOrders() {
  const cfg = await mpIdb("readonly", st => st.get("cfg"));
  if (!cfg || !cfg.token || !cfg.api) return;
  const ymd = d => d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
  const from = new Date(Date.now() - 864e5), to = new Date(Date.now() + 60 * 864e5);
  const r = await fetch(cfg.api, { method: "POST", redirect: "follow",
    body: JSON.stringify({ action: "adminBookings", token: cfg.token, from: ymd(from), to: ymd(to) }) });
  const j = await r.json(); if (!j.ok) return;
  const seen = new Set(cfg.seen || []);
  const fresh = j.bookings.filter(b => b.bookingid && !seen.has(String(b.bookingid)) &&
    String(b.source || "") !== "Admin" && !/^(cancelled|no-show)$/i.test(String(b.status || "")));
  if (!fresh.length) return;
  fresh.forEach(b => seen.add(String(b.bookingid)));
  cfg.seen = Array.from(seen).slice(-1500);
  await mpIdb("readwrite", st => st.put(cfg, "cfg"));
  const open = await clients.matchAll({ type: "window" });
  if (open.some(c => c.visibilityState === "visible")) return;   // the app is on screen and will ring by itself
  const b = fresh[0];
  await self.registration.showNotification(fresh.length > 1 ? fresh.length + " new orders" : "New order · " + (b.customername || "Customer"), {
    body: fresh.length > 1 ? fresh.slice(0, 4).map(x => (x.customername || "Customer") + " · " + (x.servicename || "")).join("\n") : [b.servicename, b.date, b.time].filter(Boolean).join(" · "),
    tag: "mp-new-order", renotify: true, requireInteraction: true, vibrate: [450, 150, 450, 150, 450],
    icon: "apple-touch-icon.png", badge: "favicon-32.png"
  });
}
self.addEventListener("periodicsync", e => { if (e.tag === "mp-orders") e.waitUntil(mpCheckOrders().catch(() => {})); });
