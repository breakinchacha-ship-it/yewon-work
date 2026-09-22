// 앱 껍데기(화면)만 폰에 저장해 두어 인터넷이 잠깐 끊겨도 열리게 합니다. 기록 자체는 구글 시트로 갑니다.
const CACHE = 'yewon-work-v27';
const SHELL = ['./', './index.html', './manifest.webmanifest', './icon-192.png', './icon-512.png', './icon-maskable-512.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (url.origin !== location.origin) return;                       // 구글 시트·폰트는 그대로 통과
  // 화면 파일은 인터넷 우선(수정 반영), 안 되면 저장본
  e.respondWith(fetch(e.request).then(r => { const copy = r.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); return r; })
                                .catch(() => caches.match(e.request).then(m => m || caches.match('./index.html'))));
});
