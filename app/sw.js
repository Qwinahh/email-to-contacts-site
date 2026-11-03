/* Email to Contacts web app: keeps the page working offline and lets Android install it.
   Caches only the app's own files; never sees what people paste (that stays in the page). */
const CACHE = 'email-to-contacts-v2';
const FILES = ['./', 'index.html', 'app.js', 'contact.js', 'parser.js', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png', 'icon-180.png'];

self.addEventListener('install', (event) => {
  event.waitUntil(caches.open(CACHE).then((c) => c.addAll(FILES)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (event) => {
  event.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});

// Network first, so updates arrive straight away; the cache is the offline fallback. Shared text
// arrives in the query string, so match the page while ignoring it.
self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET' || new URL(event.request.url).origin !== location.origin) return;
  event.respondWith(fetch(event.request).catch(() => caches.match(event.request, { ignoreSearch: true })));
});
