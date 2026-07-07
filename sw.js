const CACHE_NAME = 'family-finance-v45';
const ASSETS = [
  './',
  './index.html',
  './styles.css',
  './js/mockData.js',
  './js/charts.js',
  './js/app.js',
  './manifest.json',
  './icon-192.png',
  './icon-512.png'
];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      console.log('Pre-caching core assets...');
      return cache.addAll(ASSETS);
    })
  );
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => {
      return Promise.all(
        keys.map(key => {
          if (key !== CACHE_NAME) {
            console.log('Deleting old cache version:', key);
            return caches.delete(key);
          }
        })
      );
    })
  );
  self.clients.claim();
});

// Stale-While-Revalidate strategy for local assets, Network-First for external sheets api
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  
  if (url.origin === self.location.origin) {
    e.respondWith(
      caches.match(e.request).then(cachedResponse => {
        if (cachedResponse) {
          // Fetch new version in background to update cache dynamically
          fetch(e.request).then(networkResponse => {
            if (networkResponse.status === 200) {
              const responseClone = networkResponse.clone();
              caches.open(CACHE_NAME).then(cache => {
                cache.put(e.request, responseClone);
              });
            }
          }).catch(() => {});
          
          return cachedResponse;
        }
        
        return fetch(e.request).then(response => {
          if (response && response.status === 200 && response.type === 'basic') {
            const responseClone = response.clone();
            caches.open(CACHE_NAME).then(cache => {
              cache.put(e.request, responseClone);
            });
          }
          return response;
        });
      })
    );
  } else {
    // External Network-First
    e.respondWith(
      fetch(e.request).then(response => {
        return response;
      }).catch(() => {
        return caches.match(e.request);
      })
    );
  }
});
