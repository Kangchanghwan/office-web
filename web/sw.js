// [cs] office.clipboardapp.org used to run an earlier editor (ranuts/document) that
// registered /sw.js. A browser that still has that worker would keep serving the old
// cached app, so this worker only clears caches and unregisters itself.
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', (event) => {
  event.waitUntil((async () => {
    for (const key of await caches.keys()) {
      if (key.startsWith('oo-x2t') || key.startsWith('document_editor')) continue; // this app's own caches
      await caches.delete(key);
    }
    await self.registration.unregister();
  })());
});
