const CACHE_NAME = 'al-majd-attendance-v1';

self.addEventListener('install', event => {
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(self.clients.claim());
});

self.addEventListener('fetch', event => {
  // لا نعترض طلبات Google Apps Script حاليًا.
});
