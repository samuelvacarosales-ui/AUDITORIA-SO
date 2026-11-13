/* Service worker: primero red (así las actualizaciones llegan solas),
   y si no hay cobertura, usa la copia guardada. */
const CACHE = 'anexo-so-v5';
const ASSETS = ['./','./index.html','./app.js','./core.js','./manifest.json',
  './xlsx.full.min.js','./jspdf.umd.min.js','./icon-192.png','./icon-512.png'];

self.addEventListener('install', e=>{
  e.waitUntil(
    caches.open(CACHE).then(c=>Promise.all(ASSETS.map(u=>c.add(u).catch(()=>null)))).then(()=>self.skipWaiting())
  );
});
self.addEventListener('activate', e=>{
  e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim()));
});
self.addEventListener('fetch', e=>{
  const r = e.request;
  if(r.method!=='GET' || new URL(r.url).origin!==location.origin) return;
  e.respondWith(
    fetch(r).then(resp=>{
      if(resp && resp.ok){ const copy=resp.clone(); caches.open(CACHE).then(c=>c.put(r,copy)); }
      return resp;
    }).catch(()=>caches.match(r).then(m=>m||caches.match('./index.html')))
  );
});
