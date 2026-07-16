const CACHE = 'drip-v1.0.0';

// Static assets to precache on install
const PRECACHE = [
  './farm-irrigation-planner.html',
  './soil-raster.png',
  'https://unpkg.com/leaflet@1.9.4/dist/leaflet.css',
  'https://unpkg.com/leaflet@1.9.4/dist/leaflet.js',
  'https://cdnjs.cloudflare.com/ajax/libs/leaflet.draw/1.0.4/leaflet.draw.css',
  'https://cdnjs.cloudflare.com/ajax/libs/leaflet.draw/1.0.4/leaflet.draw.js',
  'https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js',
];

// Domains that should always go to the network (weather, map tiles, AI)
const NETWORK_ONLY = [
  'api.open-meteo.com',
  'arcgisonline.com',
  'nominatim.openstreetmap.org',
  'ai-agent.andrespatrignani.workers.dev',
];

self.addEventListener('install', e=>{
  e.waitUntil(
    caches.open(CACHE)
      .then(c=>c.addAll(PRECACHE))
      .then(()=>self.skipWaiting())
  );
});

self.addEventListener('activate', e=>{
  e.waitUntil(
    caches.keys()
      .then(keys=>Promise.all(keys.filter(k=>k!==CACHE).map(k=>caches.delete(k))))
      .then(()=>self.clients.claim())
  );
});

self.addEventListener('fetch', e=>{
  const url = new URL(e.request.url);

  // Always network for dynamic/live endpoints
  if(NETWORK_ONLY.some(d=>url.hostname.includes(d))) return;

  // Cache-first for everything else
  e.respondWith(
    caches.match(e.request).then(cached=>{
      if(cached) return cached;
      return fetch(e.request).then(response=>{
        // Only cache successful same-origin or CORS responses
        if(response && response.status===200 && (response.type==='basic'||response.type==='cors')){
          const clone = response.clone();
          caches.open(CACHE).then(c=>c.put(e.request, clone));
        }
        return response;
      });
    })
  );
});
