const CACHE = 'cleaning-log-v19';
const ASSETS = ['./', './index.html', './styles.css?v=6', './app.js?v=19', './cloud-config.js', './manifest.webmanifest', './icon.svg', './icon-192.png', './icon-512.png', './vendor/docx.iife.js', './supabase.min.js'];
self.addEventListener('install', (event) => event.waitUntil(caches.open(CACHE).then((cache) => cache.addAll(ASSETS)).then(() => self.skipWaiting())));
self.addEventListener('activate', (event) => event.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((key) => key !== CACHE).map((key) => caches.delete(key)))).then(() => self.clients.claim())));
self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;
  if (event.request.mode === 'navigate') {
    event.respondWith(fetch(event.request).then((response) => { const copy = response.clone(); caches.open(CACHE).then((cache) => cache.put('./', copy)); return response; }).catch(() => caches.match('./')));
    return;
  }
  event.respondWith(caches.match(event.request).then((cached) => {
    const fresh = fetch(event.request).then((response) => { if (response.ok && new URL(event.request.url).origin === self.location.origin) { const copy = response.clone(); caches.open(CACHE).then((cache) => cache.put(event.request, copy)); } return response; });
    return cached || fresh;
  }));
});
