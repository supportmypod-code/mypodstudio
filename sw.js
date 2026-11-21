/* MyPod service worker. Bump VERSION on every release so customers get the new files. */
const VERSION = "mypod-v13";
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


/* Tapping a notification opens the site at "My booking". */
self.addEventListener("notificationclick", e => {
  e.notification.close();
  e.waitUntil((async () => {
    const all = await clients.matchAll({ type: "window", includeUncontrolled: true });
    for (const c of all) { if ("focus" in c) { await c.focus(); return; } }
    if (clients.openWindow) await clients.openWindow("./#manage");
  })());
});

/* Background check of the customer's bookings (Chrome / Edge on Android, installed app only). */
function mpcIdb(mode, fn) {
  return new Promise((res, rej) => {
    const r = indexedDB.open("mypodcust", 1);
    r.onupgradeneeded = () => r.result.createObjectStore("kv");
    r.onerror = () => rej(r.error);
    r.onsuccess = () => {
      const tx = r.result.transaction("kv", mode), q = fn(tx.objectStore("kv"));
      tx.oncomplete = () => res(q && q.result);
      tx.onerror = () => rej(tx.error);
    };
  });
}
const MPC_TXT = {
  "In Progress": ["Your wash has started", "Our team has started on your vehicle."],
  "Completed": ["Your vehicle is ready", "Your service is complete. Thank you for choosing MyPod!"],
  "Cancelled": ["Booking cancelled", "Your booking has been cancelled."],
  "Confirmed": ["Booking confirmed", "Your booking is confirmed."],
  "No-show": ["Booking closed", "This booking was marked as a no-show."]
};
async function mpcCheck() {
  const cfg = await mpcIdb("readonly", st => st.get("cfg"));
  if (!cfg || !cfg.api || !(cfg.watch || []).length) return;
  const FINAL = /^(completed|cancelled|no-show)$/i;
  let changed = false;
  for (const w of cfg.watch) {
    if (FINAL.test(w.status || "")) continue;
    try {
      const r = await fetch(cfg.api, { method: "POST", redirect: "follow",
        body: JSON.stringify({ action: "lookupBooking", bookingId: w.id, customerPhone: w.phone, website: "" }) });
      const j = await r.json(); if (!j.ok || !j.booking) continue;
      if (j.booking.status !== w.status) {
        w.status = j.booking.status; changed = true;
        const t = MPC_TXT[w.status] || ["Booking update", "Status: " + w.status];
        await self.registration.showNotification(t[0], { body: t[1], tag: "bk-" + w.id, renotify: true,
          icon: "icons/icon-192.png", badge: "icons/icon-192.png", vibrate: [200, 100, 200] });
      }
    } catch (e) {}
  }
  if (changed) await mpcIdb("readwrite", st => st.put(cfg, "cfg"));
}
self.addEventListener("periodicsync", e => { if (e.tag === "mp-booking") e.waitUntil(mpcCheck().catch(() => {})); });
