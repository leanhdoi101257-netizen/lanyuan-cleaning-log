const CACHE = 'cleaning-log-v14';
const ASSETS = ['./', './index.html', './styles.css?v=6', './app.js?v=13', './cloud-config.js', './manifest.webmanifest', './icon.svg', './icon-192.png', './icon-512.png', './vendor/docx.iife.js'];
self.addEventListener('install', (event) => event.waitUntil(caches.open(CACHE).then((cache) => cache.addAll(ASSETS)).then(() => self.skipWaiting())));
self.addEventListener('activate', (event) => event.waitUntil(self.clients.claim()));
self.addEventListener('fetch', (event) => event.respondWith(caches.match(event.request).then((cached) => cached || fetch(event.request))));
