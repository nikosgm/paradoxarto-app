// Παραδόχαρτο service worker — pdx-shell-v3 (2026-11-02)
// Κέλυφος: πάντα πρώτα το δίκτυο (για να μην κολλάει παλιά έκδοση), με αντίγραφο στην cache για χωρίς σήμα.
// Μόνο τα δικά μας αρχεία· Supabase/CDN περνούν απευθείας. Web Push: {title, body} από τη send-reminders.
const CACHE = 'pdx-shell-v7-ct';
const SHELL = /\/(index\.html|ct\.html|ct|manifest\.json|ct-manifest\.json|icon-\d+\.png|ct-icon-\d+\.png)?$/;
// 7/11: ο server κάνει ct.html → /ct (307). Μια απάντηση «με ανακατεύθυνση» ΔΕΝ επιτρέπεται για άνοιγμα σελίδας (σφάλμα «δεν είναι δυνατή η πρόσβαση»)
// → κρατάμε/δίνουμε πάντα «καθαρό» αντίγραφο.
const clean = async (r) => (r && r.redirected) ? new Response(await r.blob(), { status: r.status, statusText: r.statusText, headers: r.headers }) : r;
const VCACHE = 'pdx-vendor-v1';   // χωριστή cache: δεν σβήνεται όταν αλλάζει το κέλυφος
const VENDOR = /\/vendor\/[\w.-]+-\d+\.\d+\.\d+(\.min)?\.js$/;

self.addEventListener('install', (e) => {
  // 7/11: «Αξονικές» ως εφαρμογή υπολογιστή → ανοίγει και χωρίς internet (CD/φάκελοι)
  e.waitUntil((async () => {
    try { const c = await caches.open(CACHE); for (const u of ['ct', 'ct-manifest.json', 'ct-icon-192.png', 'ct-icon-512.png']) { try { const r = await clean(await fetch(u, { cache: 'no-store' })); if (r && r.ok) await c.put(new URL(u, self.location.href).href, r); } catch (er) {} } } catch (err) {}
    try { const v = await caches.open(VCACHE); if (!(await v.match(new URL('vendor/daikon-1.2.42.min.js', self.location.href).href))) await v.add('vendor/daikon-1.2.42.min.js'); } catch (err) {}
    self.skipWaiting();
  })());
});
self.addEventListener('activate', (e) => {
  e.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter((k) => k.startsWith('pdx-shell') && k !== CACHE).map((k) => caches.delete(k)));
    await self.clients.claim();
  })());
});
self.addEventListener('message', (e) => { if (e.data && (e.data === 'SKIP_WAITING' || e.data.type === 'SKIP_WAITING')) self.skipWaiting(); });

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  // βιβλιοθήκες με έκδοση στο όνομα (vendor/daikon-1.2.42.min.js): δεν αλλάζουν ποτέ → πρώτα η cache (χωρίς σήμα / PC νοσοκομείου)
  if (url.origin === self.location.origin && VENDOR.test(url.pathname)) {
    e.respondWith((async () => {
      const cache = await caches.open(VCACHE);
      const hit = await cache.match(url.origin + url.pathname);
      if (hit) return hit;
      const r = await fetch(req);
      if (r && r.ok) await cache.put(url.origin + url.pathname, r.clone()).catch(() => {});
      return r;
    })());
    return;
  }
  if (url.origin !== self.location.origin || !SHELL.test(url.pathname)) return;
  // ?v=… = ζητήθηκε ρητά νέα έκδοση: περιμένουμε το δίκτυο περισσότερο
  const waitMs = url.searchParams.has('v') ? 12000 : 4000;
  const key = url.origin + url.pathname;   // ένα αντίγραφο ανά αρχείο, χωρίς παραμέτρους
  e.respondWith((async () => {
    const cache = await caches.open(CACHE);
    const net = fetch(req.url, { cache: 'no-store', credentials: 'same-origin', redirect: 'follow' }).then(clean).then(async (r) => {
      if (r && r.ok) await cache.put(key, r.clone()).catch(() => {});
      return r;
    });
    const timeout = new Promise((res) => setTimeout(() => res(null), waitMs));
    try {
      const r = await Promise.race([net, timeout]);
      if (r && r.ok) return r;
    } catch (err) {}
    const hit = await clean(await cache.match(key));
    if (hit) { e.waitUntil(net.catch(() => {})); return hit; }
    return net;   // καμία cache: περιμένουμε όσο χρειαστεί
  })());
});

self.addEventListener('push', (e) => {
  let d = {};
  try { d = e.data ? e.data.json() : {}; } catch (err) { d = { body: e.data ? e.data.text() : '' }; }
  const title = d.title || 'Παραδόχαρτο';
  e.waitUntil(self.registration.showNotification(title, {
    body: d.body || '',
    icon: 'icon-192.png',
    badge: 'icon-192.png',
    tag: d.tag || undefined,
    data: { url: d.url || './', title, kind: /για (έλεγχο|έγκριση)/.test(title + ' ' + (d.body || '')) ? 'review'
      // αποτυχίες/κολλήματα εργασιών bot → σελίδα «Εργασίες bot»
      : /⚠|δεν ξεκίνησε|απέτυχε|υπολογιστ|φύλακας/.test(title + ' ' + (d.body || '')) && !/εκτός ορίων|κωδικ/i.test(title + ' ' + (d.body || '')) ? 'jobs'
      // 9/11: εργαστηριακά / καλλιέργειες / κρίσιμες τιμές → σελίδα 🔔 Ειδοποιήσεις (πριν: απλώς άνοιγε η εφαρμογή)
      : /εκτός ορίων|καλλιέργ|κρίσιμ|🚨|εργαστηρ/i.test(title + ' ' + (d.body || '')) ? 'notif' : '' },
  }));
});

self.addEventListener('notificationclick', (e) => {
  e.notification.close();
  const nd = e.notification.data || {}, rev = nd.kind === 'review';
  // «για έλεγχο/έγκριση» → η εφαρμογή ανοίγει κατευθείαν το φύλλο ελέγχου
  const jobs = nd.kind === 'jobs', nt = nd.kind === 'notif';
  const tu = new URL(nd.url || './', self.location.href); if (rev) tu.hash = 'pdxrev'; else if (jobs) tu.hash = 'pdxjobs'; else if (nt) tu.hash = 'pdxnotif'; const target = tu.href;
  e.waitUntil((async () => {
    const all = await self.clients.matchAll({ type: 'window', includeUncontrolled: true });
    for (const c of all) { if (c.url.startsWith(self.location.origin)) { await c.focus(); if (rev) c.postMessage({ type: 'pdx-open-review' }); else if (jobs) c.postMessage({ type: 'pdx-open-jobs' }); else if (nt) c.postMessage({ type: 'pdx-open-notif' }); return; } }
    await self.clients.openWindow(target);
  })());
});
