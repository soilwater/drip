// Bump this on every release. It is the ONLY thing that tells browsers a new
// build exists — the old cache is deleted and clients reload automatically.
// Keep in sync with APP_VERSION in index.html.
const CACHE = 'drip-v1.2.0';

// App shell. Every path here must resolve, or addAll() rejects and the whole
// install fails silently — leaving users with no service worker at all.
const PRECACHE = [
  './',
  './index.html',
  './manifest.json',
  './soil_texture.png',
  './icon.svg',
  './icon-180.png',
  './icon-192.png',
  './icon-512.png',
  'https://unpkg.com/leaflet@1.9.4/dist/leaflet.css',
  'https://unpkg.com/leaflet@1.9.4/dist/leaflet.js',
  'https://cdnjs.cloudflare.com/ajax/libs/leaflet.draw/1.0.4/leaflet.draw.css',
  'https://cdnjs.cloudflare.com/ajax/libs/leaflet.draw/1.0.4/leaflet.draw.js',
  'https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js',
];

// Live data — never served from cache.
const NETWORK_ONLY = [
  'api.open-meteo.com',
  'archive-api.open-meteo.com',
  'arcgisonline.com',
  'ai-agent.andrespatrignani.workers.dev',
];

self.addEventListener('install', e=>{
  e.waitUntil((async ()=>{
    const c = await caches.open(CACHE);
    // Cache entries individually so one bad URL (e.g. a CDN hiccup) can't fail
    // the whole install the way addAll() would.
    await Promise.all(PRECACHE.map(url=>
      c.add(new Request(url, {cache:'reload'})).catch(err=>console.warn('precache miss:', url, err))
    ));
    // Deliberately NOT calling skipWaiting() here. The new build sits in the
    // "waiting" state so the page can offer a Refresh prompt instead of pulling
    // the rug out from under someone mid-form.
  })());
});

// The page asks us to take over once the user taps Refresh.
self.addEventListener('message', e=>{
  if(e.data && e.data.type==='SKIP_WAITING') self.skipWaiting();
});

self.addEventListener('activate', e=>{
  e.waitUntil((async ()=>{
    const keys = await caches.keys();
    await Promise.all(keys.filter(k=>k!==CACHE).map(k=>caches.delete(k)));
    await self.clients.claim();  // take over open pages -> triggers controllerchange
  })());
});

self.addEventListener('fetch', e=>{
  const req = e.request;
  if(req.method !== 'GET') return;
  const url = new URL(req.url);
  if(NETWORK_ONLY.some(d=>url.hostname.includes(d))) return;

  // Navigations: network-first, so a fresh index.html is picked up as soon as
  // there's a connection, with the cached shell as an offline fallback.
  if(req.mode === 'navigate'){
    e.respondWith(
      fetch(req)
        .then(res=>{
          const clone = res.clone();
          caches.open(CACHE).then(c=>c.put('./index.html', clone));
          return res;
        })
        .catch(()=>caches.match('./index.html'))
    );
    return;
  }

  // Everything else: cache-first (libraries and the soil raster are versioned
  // or static, so this is safe and keeps startup fast).
  e.respondWith(
    caches.match(req).then(cached=>
      cached || fetch(req).then(res=>{
        if(res && res.status===200 && (res.type==='basic'||res.type==='cors')){
          const clone = res.clone();
          caches.open(CACHE).then(c=>c.put(req, clone));
        }
        return res;
      })
    )
  );
});
