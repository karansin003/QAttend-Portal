const CACHE_VERSION = "qattend-v1";
const APP_SHELL = [
    "./",
    "./index.html",
    "./login.html",
    "./admin.html",
    "./cr.html",
    "./style.css",
    "./script.js",
    "./logo.png"
];

// Firebase/SheetJS/EmailJS network calls are deliberately NOT cached here.
// Firestore has its own IndexedDB offline cache and write queue.
const EXTERNAL_ASSETS = [
    "https://www.gstatic.com/firebasejs/12.18.0/firebase-app.js",
    "https://www.gstatic.com/firebasejs/12.18.0/firebase-auth.js",
    "https://www.gstatic.com/firebasejs/12.18.0/firebase-firestore.js",
    "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js"
];

self.addEventListener("install", event => {
    event.waitUntil(
        caches.open(CACHE_VERSION).then(async cache => {
            // Cache local app shell individually so one missing optional asset
            // (for example logo.png in an older deployment) cannot abort install.
            await Promise.allSettled(
                APP_SHELL.map(async url => {
                    try {
                        await cache.add(url);
                    } catch (error) {
                        console.warn("Could not cache app asset:", url, error);
                    }
                })
            );

            // Firebase modules are needed to bootstrap the app offline after the
            // first successful online visit. Cache them when the CDN allows it.
            await Promise.allSettled(
                EXTERNAL_ASSETS.map(async url => {
                    try {
                        const response = await fetch(url, { mode: "cors" });
                        if (response.ok) await cache.put(url, response);
                    } catch (error) {
                        console.warn("Could not cache external asset:", url, error);
                    }
                })
            );
        }).then(() => self.skipWaiting())
    );
});

self.addEventListener("activate", event => {
    event.waitUntil(
        caches.keys().then(keys =>
            Promise.all(
                keys
                    .filter(key => key !== CACHE_VERSION)
                    .map(key => caches.delete(key))
            )
        ).then(() => self.clients.claim())
    );
});

self.addEventListener("message", event => {
    if (event.data?.type === "SKIP_WAITING") {
        self.skipWaiting();
    }
});

self.addEventListener("fetch", event => {
    const request = event.request;
    if (request.method !== "GET") return;

    const url = new URL(request.url);

    // Never intercept Firebase backend APIs. Firestore/Auth handle their own
    // offline behavior; Service Worker caching those requests can cause stale
    // or incorrect authentication/data behavior.
    const isFirebaseBackend =
        url.hostname.endsWith("googleapis.com") ||
        url.hostname.endsWith("firebaseio.com");

    if (isFirebaseBackend) return;

    // Navigation: network first when available, cached page when offline.
    if (request.mode === "navigate") {
        event.respondWith(
            fetch(request)
                .then(response => {
                    const copy = response.clone();
                    caches.open(CACHE_VERSION).then(cache => cache.put(request, copy));
                    return response;
                })
                .catch(() =>
                    caches.match(request).then(cached =>
                        cached || caches.match("./login.html")
                    )
                )
        );
        return;
    }

    // Same-origin static files: cache first for fast repeat/offline loads.
    if (url.origin === self.location.origin) {
        event.respondWith(
            caches.match(request).then(cached => {
                if (cached) return cached;

                return fetch(request).then(response => {
                    if (response.ok) {
                        const copy = response.clone();
                        caches.open(CACHE_VERSION).then(cache => cache.put(request, copy));
                    }
                    return response;
                });
            })
        );
        return;
    }

    // Known CDN modules: cache first, then network.
    if (
        url.hostname === "www.gstatic.com" ||
        url.hostname === "cdn.sheetjs.com"
    ) {
        event.respondWith(
            caches.match(request).then(cached =>
                cached || fetch(request).then(response => {
                    if (response.ok) {
                        const copy = response.clone();
                        caches.open(CACHE_VERSION).then(cache => cache.put(request, copy));
                    }
                    return response;
                })
            )
        );
    }
});
