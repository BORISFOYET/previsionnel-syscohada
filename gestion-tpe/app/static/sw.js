// Service worker minimal : rend le site installable et garde en cache la charte (pas les données).
const C = 'tpe-v1';
self.addEventListener('install', e => e.waitUntil(caches.open(C).then(c => c.addAll(['/static/style.css', '/static/logo_accounts.jpg']))));
self.addEventListener('fetch', e => {
  if (e.request.method === 'GET' && new URL(e.request.url).pathname.startsWith('/static/'))
    e.respondWith(caches.match(e.request).then(r => r || fetch(e.request)));
});
