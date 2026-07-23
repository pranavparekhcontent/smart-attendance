/**
 * VibeMantra Firebase Push Notification Configuration
 * Shared across: Smart Attendance, Academic File App, Student Portal
 */
const FIREBASE_CONFIG = {
  apiKey: "AIzaSyBuw7HMI__3oNgMbjQz-q2L1aoIcfn5H9k",
  authDomain: "vibemantra-pwa-f3453.firebaseapp.com",
  projectId: "vibemantra-pwa-f3453",
  storageBucket: "vibemantra-pwa-f3453.firebasestorage.app",
  messagingSenderId: "1022530251175",
  appId: "1:1022530251175:web:237a07fb1eb83c87fd6e82",
  measurementId: "G-544HRXQQ33"
};

const VAPID_KEY = "BGKXij5LeLw-DFR8NWSGmvWl36xMfLJYcQTntQTrPbVviX28ApXhAY-8DkcefT7EU8qSomQeNDWb6dHGYAbBdqs";

if (typeof window !== 'undefined') {
  window.FIREBASE_CONFIG = FIREBASE_CONFIG;
  window.VAPID_KEY = VAPID_KEY;
}
