// 陈平安资料库 Service Worker
const CACHE_NAME = 'cpa-db-v1';
const STATIC_ASSETS = [
    '/',
    '/static/css/style.css',
    '/static/manifest.json',
];

// 安装：缓存静态资源
self.addEventListener('install', function(event) {
    self.skipWaiting();
});

// 激活：清理旧缓存
self.addEventListener('activate', function(event) {
    event.waitUntil(
        caches.keys().then(function(names) {
            return Promise.all(
                names.filter(function(name) { return name !== CACHE_NAME; })
                     .map(function(name) { return caches.delete(name); })
            );
        })
    );
    self.clients.claim();
});

// 请求策略：网络优先，失败回退缓存
self.addEventListener('fetch', function(event) {
    if (event.request.method !== 'GET') return;

    event.respondWith(
        fetch(event.request)
            .then(function(response) {
                // 成功则缓存副本
                if (response.status === 200) {
                    var responseClone = response.clone();
                    caches.open(CACHE_NAME).then(function(cache) {
                        cache.put(event.request, responseClone);
                    });
                }
                return response;
            })
            .catch(function() {
                // 网络失败，尝试缓存
                return caches.match(event.request);
            })
    );
});
