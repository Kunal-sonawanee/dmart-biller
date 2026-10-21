// Offline support: the app keeps working without internet once it has been opened.
// Bump CACHE_VERSION whenever styles.css or vendor files change.
const CACHE_VERSION = 'dmart-bills-v2';
const APP_SHELL = [
    './',
    './styles.css',
    './manifest.json',
    './vendor/html2pdf.bundle.min.js',
    './vendor/xlsx.full.min.js',
    './icons/icon-192.png',
    './icons/icon-512.png',
    './icons/apple-touch-icon.png'
];

self.addEventListener('install', event => {
    event.waitUntil(caches.open(CACHE_VERSION).then(cache => cache.addAll(APP_SHELL)));
    self.skipWaiting();
});

self.addEventListener('activate', event => {
    event.waitUntil(
        caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE_VERSION).map(k => caches.delete(k))))
    );
    self.clients.claim();
});

self.addEventListener('fetch', event => {
    const req = event.request;
    if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;

    // The page itself: try the network first so updates show up, fall back to the cached copy offline.
    if (req.mode === 'navigate') {
        event.respondWith(
            fetch(req)
                .then(res => {
                    // Hosts may redirect /index.html to /; a redirected response can't be replayed offline
                    if (res.ok && !res.redirected) {
                        const copy = res.clone();
                        caches.open(CACHE_VERSION).then(cache => cache.put('./', copy));
                    }
                    return res;
                })
                .catch(() => caches.match('./'))
        );
        return;
    }

    // Everything else: cached copy first.
    event.respondWith(caches.match(req).then(hit => hit || fetch(req)));
});
