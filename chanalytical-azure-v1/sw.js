const CACHE = 'cha-v20';

// Critical files — must cache for offline to work
const CRITICAL = [
  '/',
  '/index.html',
  '/login.html',
  '/dashboard.html',
  '/auth.js',
  '/manifest.json',
  '/offline.html',
];

// Nice-to-have — cache if possible but don't fail install if they don't load
const OPTIONAL = [
  '/admin-dashboard.html',
  '/icon-192.png',
  '/icon-512.png',
  'https://cdnjs.cloudflare.com/ajax/libs/jsbarcode/3.11.5/JsBarcode.all.min.js',
];

self.addEventListener('install', event => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE);

    // Cache critical files — if ANY of these fail, log but continue
    // (iOS Safari will still activate the SW)
    for (const url of CRITICAL) {
      try {
        await cache.add(url);
      } catch (e) {
        console.warn('[SW] Critical cache miss:', url, e.message);
      }
    }

    // Cache optional files — failures are silent
    for (const url of OPTIONAL) {
      try { await cache.add(url); } catch {}
    }

    // Skip waiting — become active immediately without waiting for old tabs to close
    await self.skipWaiting();
  })())
});

self.addEventListener('activate', event => {
  event.waitUntil((async () => {
    // Delete all old caches
    const keys = await caches.keys();
    await Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)));
    // Take control of all open pages immediately
    await self.clients.claim();
  })())
});

self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  // Only handle GET requests over http/https
  if (event.request.method !== 'GET') return;
  if (!url.protocol.startsWith('http')) return;

  // Netlify functions — always try network, return JSON error if offline
  if (url.pathname.startsWith('/.netlify/')) {
    event.respondWith(
      fetch(event.request).catch(() =>
        new Response(JSON.stringify({ offline: true }), {
          status: 503,
          headers: { 'Content-Type': 'application/json' }
        })
      )
    );
    return;
  }

  // Everything else: cache-first
  event.respondWith((async () => {
    // 1. Check cache first
    const cached = await caches.match(event.request, { ignoreSearch: true });
    if (cached) {
      // Update cache in background while serving cached response
      fetch(event.request).then(r => {
        if (r && r.status === 200) {
          caches.open(CACHE).then(c => c.put(event.request, r));
        }
      }).catch(() => {});
      return cached;
    }

    // 2. Try network
    try {
      const response = await fetch(event.request);
      if (response && response.status === 200) {
        const clone = response.clone();
        caches.open(CACHE).then(c => c.put(event.request, clone));
      }
      return response;
    } catch {
      // 3. Network failed — serve fallback
      if (event.request.mode === 'navigate') {
        // Try to serve the main app page from cache
        return (
          await caches.match('/index.html') ||
          await caches.match('/') ||
          await caches.match('/offline.html') ||
          new Response('<h1>Offline</h1><p>Please connect to internet and reload once to enable offline mode.</p>', {
            headers: { 'Content-Type': 'text/html' }
          })
        );
      }
      return new Response('', { status: 408 });
    }
  })());
});

// Background sync (Android Chrome only — iOS ignores this)
self.addEventListener('sync', event => {
  if (event.tag === 'cha-sync-queue') {
    event.waitUntil(bgSync());
  }
});

async function bgSync() {
  const db = await openIDB().catch(() => null);
  if (!db) return;

  const queue = await getAllFromStore(db, 'queue').catch(() => []);
  const unsynced = queue.filter(s => !s.synced);

  for (const sub of unsynced) {
    try {
      const r = await fetch('/.netlify/functions/sync-to-sheets', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(sub),
      });
      if (r.ok) {
        sub.synced = true;
        sub.syncedAt = new Date().toISOString();
        await putInStore(db, 'queue', sub).catch(() => {});
      }
    } catch {}
  }

  const clients = await self.clients.matchAll({ type: 'window' });
  clients.forEach(c => c.postMessage({ type: 'SYNC_COMPLETE' }));
}

// IDB helpers
function openIDB() {
  return new Promise((res, rej) => {
    const r = indexedDB.open('cha_db', 1);
    r.onupgradeneeded = e => e.target.result.createObjectStore('queue', { keyPath: 'id' });
    r.onsuccess = e => res(e.target.result);
    r.onerror = () => rej(r.error);
  });
}

function getAllFromStore(db, store) {
  return new Promise((res, rej) => {
    const r = db.transaction(store, 'readonly').objectStore(store).getAll();
    r.onsuccess = () => res(r.result || []);
    r.onerror = () => rej(r.error);
  });
}

function putInStore(db, store, item) {
  return new Promise((res, rej) => {
    const r = db.transaction(store, 'readwrite').objectStore(store).put(item);
    r.onsuccess = res;
    r.onerror = () => rej(r.error);
  });
}
