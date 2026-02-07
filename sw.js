/**
 * 용인대학교 출석시스템 Service Worker
 * Version: 1.0.0
 *
 * 캐싱 전략: Network First (네트워크 우선, 실패 시 캐시)
 * - 출석 시스템은 실시간 데이터가 중요하므로 네트워크 우선
 * - 오프라인 시 기본 페이지 제공
 */

const CACHE_NAME = 'yiu-attendance-v1';
const OFFLINE_URL = 'offline.html';

// 캐시할 정적 리소스
const STATIC_ASSETS = [
  './',
  './index.html',
  './admin.html',
  './offline.html',
  './manifest.json',
  './icons/icon-192.png',
  './icons/icon-512.png'
];

/**
 * Service Worker 설치
 * 정적 리소스를 캐시에 저장
 */
self.addEventListener('install', (event) => {
  console.log('[SW] 설치 중...');

  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => {
        console.log('[SW] 정적 리소스 캐싱');
        // 정적 리소스만 캐시 (CDN은 런타임에 캐시)
        return cache.addAll(STATIC_ASSETS.filter(url => !url.startsWith('http')));
      })
      .then(() => {
        // 즉시 활성화
        return self.skipWaiting();
      })
      .catch((error) => {
        console.error('[SW] 캐시 실패:', error);
      })
  );
});

/**
 * Service Worker 활성화
 * 이전 버전 캐시 삭제
 */
self.addEventListener('activate', (event) => {
  console.log('[SW] 활성화 중...');

  event.waitUntil(
    caches.keys()
      .then((cacheNames) => {
        return Promise.all(
          cacheNames
            .filter((name) => name !== CACHE_NAME)
            .map((name) => {
              console.log('[SW] 이전 캐시 삭제:', name);
              return caches.delete(name);
            })
        );
      })
      .then(() => {
        // 모든 클라이언트 즉시 제어
        return self.clients.claim();
      })
  );
});

/**
 * 네트워크 요청 처리
 * Network First 전략 사용
 */
self.addEventListener('fetch', (event) => {
  const { request } = event;
  const url = new URL(request.url);

  // Google Apps Script API 요청은 캐시하지 않음
  if (url.hostname.includes('script.google.com') ||
      url.hostname.includes('script.googleusercontent.com')) {
    return;
  }

  // POST 요청은 캐시하지 않음
  if (request.method !== 'GET') {
    return;
  }

  event.respondWith(
    networkFirst(request)
  );
});

/**
 * Network First 전략
 * 1. 네트워크에서 먼저 가져오기 시도
 * 2. 실패 시 캐시에서 가져오기
 * 3. 캐시도 없으면 오프라인 페이지
 */
async function networkFirst(request) {
  try {
    // 네트워크 요청 시도
    const networkResponse = await fetch(request);

    // 성공하면 캐시에 저장
    if (networkResponse.ok) {
      const cache = await caches.open(CACHE_NAME);
      cache.put(request, networkResponse.clone());
    }

    return networkResponse;
  } catch (error) {
    console.log('[SW] 네트워크 실패, 캐시에서 로드:', request.url);

    // 캐시에서 찾기
    const cachedResponse = await caches.match(request);

    if (cachedResponse) {
      return cachedResponse;
    }

    // HTML 요청이면 오프라인 페이지 반환
    if (request.headers.get('Accept')?.includes('text/html')) {
      const offlineResponse = await caches.match(OFFLINE_URL);
      if (offlineResponse) {
        return offlineResponse;
      }
    }

    // 기본 오류 응답
    return new Response('오프라인 상태입니다.', {
      status: 503,
      statusText: 'Service Unavailable',
      headers: { 'Content-Type': 'text/plain; charset=utf-8' }
    });
  }
}

/**
 * 푸시 알림 수신 (향후 확장용)
 */
self.addEventListener('push', (event) => {
  if (!event.data) return;

  const data = event.data.json();
  const options = {
    body: data.body || '새로운 알림이 있습니다.',
    icon: './icons/icon-192.png',
    badge: './icons/icon-72.png',
    vibrate: [100, 50, 100],
    data: {
      url: data.url || './'
    }
  };

  event.waitUntil(
    self.registration.showNotification(
      data.title || '용인대학교 출석시스템',
      options
    )
  );
});

/**
 * 알림 클릭 처리
 */
self.addEventListener('notificationclick', (event) => {
  event.notification.close();

  const url = event.notification.data?.url || './';

  event.waitUntil(
    clients.matchAll({ type: 'window', includeUncontrolled: true })
      .then((windowClients) => {
        // 이미 열린 창이 있으면 포커스
        for (const client of windowClients) {
          if (client.url.includes(url) && 'focus' in client) {
            return client.focus();
          }
        }
        // 없으면 새 창 열기
        if (clients.openWindow) {
          return clients.openWindow(url);
        }
      })
  );
});

console.log('[SW] Service Worker 로드됨');
