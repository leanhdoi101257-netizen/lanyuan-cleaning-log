const CACHE = 'cleaning-log-v17';
const ASSETS = ['./', './index.html', './styles.css?v=6', './app.js?v=14', './cloud-config.js', './manifest.webmanifest', './icon.svg', './icon-192.png', './icon-512.png', './vendor/docx.iife.js'];
self.addEventListener('install', (event) => event.waitUntil(caches.open(CACHE).then((cache) => cache.addAll(ASSETS)).then(() => self.skipWaiting())));
self.addEventListener('activate', (event) => event.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((key) => key !== CACHE).map((key) => caches.delete(key)))).then(() => self.clients.claim())));
self.addEventListener('fetch', (event) => event.respondWith(caches.match(event.request).then((cached) => cached || fetch(event.request))));
