import { initializeApp } from 'firebase/app';
import { getFirestore } from 'firebase/firestore';
import { getStorage } from 'firebase/storage';

const firebaseConfig = {
  apiKey: import.meta.env.VITE_FIREBASE_API_KEY,
  authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN || "logistic-automation-suite.firebaseapp.com",
  projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID || "logistic-automation-suite",
  storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET || "logistic-automation-suite.firebasestorage.app",
  messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID || "627729067946",
  appId: import.meta.env.VITE_FIREBASE_APP_ID || "1:627729067946:web:8b1675dfdcf2aee76a05d3",
  measurementId: import.meta.env.VITE_FIREBASE_MEASUREMENT_ID || "G-WB0W73XGMY",
};

const app = initializeApp(firebaseConfig);

export const db = getFirestore(app);
export const storage = getStorage(app);