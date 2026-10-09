
/* NexChat Service Worker v3
 * Private chats and account pages are intentionally NOT cached, preventing
 * stale account UI/messages from reappearing after switching accounts.
 */
const CACHE_PREFIX = "nexchat-private-safe-";
self.addEventListener("install", event => { self.skipWaiting(); });
self.addEventListener("activate", event => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter(k => k.startsWith(CACHE_PREFIX) || /viora|nexchat/i.test(k)).map(k => caches.delete(k)));
    await self.clients.claim();
  })());
});
self.addEventListener("fetch", event => {
  const req = event.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  // Always use network for the app document and dynamic/API requests; never serve a cached account.
  if (req.mode === "navigate" || /script\.google\.com|googleapis\.com|drive\.google\.com/.test(url.hostname)) {
    event.respondWith(fetch(req).catch(() => new Response("You are offline. Reconnect to open NexChat safely.", {status: 503, headers:{"Content-Type":"text/plain; charset=utf-8"}})));
  }
});