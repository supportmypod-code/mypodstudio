/* MyPod Admin service worker: caches the app shell so the app opens offline.
   Booking data and Google sign-in are never cached here. */
const VERSION = 'mypod-admin-v3';
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
