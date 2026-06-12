self.addEventListener('install', (e) => {
    self.skipWaiting();
});

self.addEventListener('activate', (e) => {
    e.waitUntil(clients.claim());
});

self.addEventListener('fetch', (e) => {
    // Carrega direto da rede para garantir o sincronismo com o Google Sheets em tempo real
    e.respondWith(fetch(e.request).catch(() => caches.match(e.request)));
});