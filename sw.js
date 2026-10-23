// Necesario para que el celular permita instalar la pagina como app.
// No guarda nada: siempre pide todo a internet para que las marcadas sean en tiempo real.
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));
self.addEventListener('fetch', e => e.respondWith(fetch(e.request)));
