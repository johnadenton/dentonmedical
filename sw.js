// dentonmedical service worker - network first, always fresh
const CACHE_NAME = 'dentonmedical-v1';

// On install - skip waiting so new SW takes over immediately
self.addEventListener('install', e => {
  self.skipWaiting();
});

// On activate - delete ALL old caches and claim clients immediately
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

// NETWORK FIRST for everything - cache is only a fallback for offline
self.addEventListener('fetch', e => {
  // Never intercept API calls or Apps Script calls
  if (e.request.url.includes('script.google.com') ||
      e.request.url.includes('api.anthropic.com') ||
      e.request.url.includes('googleapis.com')) {
    return; // Let browser handle these directly
  }

  e.respondWith(
    fetch(e.request)
      .then(res => {
        // Got network response - update cache and return it
        if (res && res.status === 200) {
          const clone = res.clone();
          caches.open(CACHE_NAME).then(c => c.put(e.request, clone));
        }
        return res;
      })
      .catch(() => {
        // Network failed - try cache as fallback
        return caches.match(e.request);
      })
  );
});
