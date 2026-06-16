/* 장원급제반 — FCM 웹 푸시 서비스워커 (백그라운드 알림 표시) */
importScripts('https://www.gstatic.com/firebasejs/12.12.1/firebase-app-compat.js');
importScripts('https://www.gstatic.com/firebasejs/12.12.1/firebase-messaging-compat.js');

firebase.initializeApp({
  apiKey:            "AIzaSyDRViuzADHxHnvv78jfnvPbAJ5hH3wzpCw",
  authDomain:        "topclass-be740.firebaseapp.com",
  projectId:         "topclass-be740",
  storageBucket:     "topclass-be740.firebasestorage.app",
  messagingSenderId: "662651378616",
  appId:             "1:662651378616:web:6eceb1505265c0700bcd4d"
});

const messaging = firebase.messaging();

messaging.onBackgroundMessage(function(payload) {
  const n = payload.notification || {};
  self.registration.showNotification(n.title || '장원급제반', {
    body:  n.body || '',
    icon:  'https://tjddlf0224-sudo.github.io/Top_Class/icon.png',
    badge: 'https://tjddlf0224-sudo.github.io/Top_Class/icon.png'
  });
});

/* 알림 클릭 시 앱 열기 */
self.addEventListener('notificationclick', function(event) {
  event.notification.close();
  event.waitUntil(clients.openWindow('https://tjddlf0224-sudo.github.io/Top_Class/'));
});
