// Lalitha Naturals - Shift Register: minimal offline-shell service worker.
//
// This is a REAL, standalone file registered at a stable relative URL
// (./sw.js) from the main app's inline registration script - not a Blob
// URL. That's what lets the browser's native service-worker update
// algorithm (byte-diffing this file against the previously-installed
// version at this same URL on every navigation / periodic check) actually
// detect new deploys and swap the worker, instead of silently never
// updating.
//
// IMPORTANT FOR MAINTAINERS (including future instances of Claude):
// CACHE_NAME below is hardcoded to the app's APP_VERSION because this file
// lives OUTSIDE the main HTML's own <script>, where APP_VERSION is a JS
// variable it can read directly. This file cannot reference that variable.
// >>> Every time APP_VERSION is bumped in the main HTML, bump the version
// >>> suffix in CACHE_NAME below to match, or the activate-time cache
// >>> cleanup logic stops doing anything useful across versions.
const CACHE_NAME = 'lalitha-register-shell-v1.2.6'; // keep in sync with APP_VERSION

self.addEventListener('install', function (e) {
    self.skipWaiting();
});

self.addEventListener('activate', function (e) {
    e.waitUntil(Promise.all([
        clients.claim(),
        caches.keys().then(function (names) {
            return Promise.all(
                names
                    .filter(function (n) { return n !== CACHE_NAME; })
                    .map(function (n) { return caches.delete(n); })
            );
        })
    ]));
});

self.addEventListener('fetch', function (event) {
    const req = event.request;
    if (req.method !== 'GET') return;

    const url = new URL(req.url);

    // Cloud sync must always be live - never intercepted or cached.
    if (
        url.hostname === 'script.google.com' ||
        url.hostname.endsWith('.script.google.com') ||
        url.hostname === 'script.googleusercontent.com'
    ) {
        return;
    }

    if (req.mode === 'navigate') {
        // Network-first for the page itself: on success, refresh the cache
        // but always serve what the network just returned, so an update is
        // visible the moment it is reopened. Only falls back to the cached
        // shell if the network request fails outright (offline).
        event.respondWith(
            fetch(req).then(function (resp) {
                const copy = resp.clone();
                caches.open(CACHE_NAME).then(function (c) { c.put(req, copy); }).catch(function () {});
                return resp;
            }).catch(function () {
                return caches.open(CACHE_NAME).then(function (c) { return c.match(req); });
            })
        );
    }
});
