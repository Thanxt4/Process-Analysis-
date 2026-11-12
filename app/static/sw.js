// Service worker: ทำให้ติดตั้งเป็นแอปบนมือถือได้ และเปิดหน้าแอปได้เร็วขึ้น
const CACHE = "pc-shell-v147";
const SHELL = ["/", "/retort.html", "/manifest.webmanifest", "/icon-192.png", "/icon-512.png"];
self.addEventListener("install", e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL))); self.skipWaiting(); });
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))));
  self.clients.claim();
});
self.addEventListener("fetch", e => {
  const url = new URL(e.request.url);
  if (e.request.method !== "GET" || url.origin !== location.origin) return;
  // /api = ข้อมูลสด, /uploads และ /t (รูปย่อ) ให้เบราว์เซอร์จำเองตาม Cache-Control (ไม่โหลดซ้ำ)
  if (url.pathname.startsWith("/api/") || url.pathname.startsWith("/uploads/") || url.pathname.startsWith("/t/") || url.pathname.startsWith("/fonts/")) return;
  // หน้า Retort: เปิดจากที่จำไว้ในเครื่องทันที (ไม่ต้องรอเน็ต) แล้วโหลดเวอร์ชันล่าสุดเก็บไว้เบื้องหลัง
  // ถ้าเวอร์ชันบนเซิร์ฟเวอร์เปลี่ยน → บอกหน้าเว็บให้ขึ้นแถบ "มีเวอร์ชันใหม่ · แตะเพื่อโหลดใหม่"
  if (url.pathname === "/retort.html") {
    const key = "/retort.html", cacheP = caches.open(CACHE), oldP = cacheP.then(c => c.match(key));
    const tag = r => r.headers.get("etag") || r.headers.get("last-modified") || "";
    const wait = ms => new Promise(r => setTimeout(r, ms));
    const tell = () => self.clients.matchAll({ type: "window" }).then(cs => cs
      .filter(c => new URL(c.url).pathname === key).forEach(c => c.postMessage({ type: "page-updated" })));
    const netP = fetch(e.request, { cache: "no-cache" }).then(async r => {
      if (r.ok) {
        const c = await cacheP, old = await oldP; await c.put(key, r.clone());
        if (old && tag(old) !== tag(r)) { await wait(1500); await tell(); await wait(3500); await tell(); }
      }
      return r;
    });
    e.waitUntil(netP.catch(() => {}));
    e.respondWith(oldP.then(old => old || netP).catch(() => netP).catch(() => Response.error()));
    return;
  }
  // network-first: ได้เวอร์ชันใหม่เสมอเมื่อออนไลน์ ถ้าเน็ตหลุดใช้ของใน cache
  // ถ้าเน็ตช้าเกิน 3 วินาที เปิดจากที่จำไว้ก่อน (เวอร์ชันใหม่จะได้ในครั้งถัดไป)
  const net = fetch(e.request, {cache: "no-cache"}).then(r => {
    if (r.ok) { const copy = r.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); }
    return r;
  });
  const fallback = () => caches.match(e.request).then(r => r || caches.match("/"));
  e.respondWith(new Promise(resolve => {
    let done = false;
    const t = setTimeout(() => fallback().then(r => { if (r && !done) { done = true; resolve(r); } }), 3000);
    net.then(r => { if (!done) { done = true; clearTimeout(t); resolve(r); } })
       .catch(() => fallback().then(r => { if (!done) { done = true; clearTimeout(t); resolve(r || Response.error()); } }));
  }));
});