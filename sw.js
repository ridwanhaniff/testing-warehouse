const CACHE_NAME = 'ina-warehouse-v4';
const ASSETS_TO_CACHE = [
    './',
    './index.html',
    './manifest.webmanifest',
    './logo.png',
    './logo-192.png',
    './logo-512.png',
    'https://cdn.tailwindcss.com',
    'https://cdn.jsdelivr.net/npm/html5-qrcode@2.3.8/html5-qrcode.min.js'
];

// Install Service Worker dan Cache File
self.addEventListener('install', (event) => {
    // skipWaiting memaksa SW baru untuk segera aktif tanpa menunggu SW lama ditutup
    event.waitUntil(
        caches.open(CACHE_NAME)
            .then((cache) => cache.addAll(ASSETS_TO_CACHE))
            .then(() => self.skipWaiting())
    );
});

// Hapus Cache Lama jika ada versi baru
self.addEventListener('activate', (event) => {
    // clients.claim memaksa SW baru untuk segera mengontrol semua halaman yang terbuka
    event.waitUntil(
        caches.keys().then((cacheNames) => {
            return Promise.all(
                cacheNames.map((cacheName) => {
                    if (cacheName !== CACHE_NAME) {
                        return caches.delete(cacheName);
                    }
                })
            );
        }).then(() => self.clients.claim())
    );
});

// Strategi: Network First, fallback to Cache
self.addEventListener('fetch', (event) => {
    if (event.request.method !== 'GET') return;

    event.respondWith(
        fetch(event.request)
            .then((response) => {
                // Jika sukses ambil dari jaringan, perbarui cache secara dinamis
                if (response && response.status === 200 && response.type === 'basic') {
                    const responseToCache = response.clone();
                    caches.open(CACHE_NAME).then((cache) => {
                        cache.put(event.request, responseToCache);
                    });
                }
                return response;
            })
            .catch(() => {
                // Jika offline atau gagal, ambil dari cache
                return caches.match(event.request);
            })
    );
});