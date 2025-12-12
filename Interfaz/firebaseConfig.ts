import { initializeApp } from 'firebase/app';
import { getFirestore } from 'firebase/firestore';
import { getStorage } from 'firebase/storage';

const firebaseConfig = {
  apiKey: "AIzaSyCXcz0Ci20GZJJ6oux88n7wCL0hcJYGk_0",
  authDomain: "logistic-automation-suite.firebaseapp.com",
  projectId: "logistic-automation-suite",
  storageBucket: "logistic-automation-suite.firebasestorage.app",
  messagingSenderId: "627729067946",
  appId: "1:627729067946:web:8b1675dfdcf2aee76a05d3",
  measurementId: "G-WB0W73XGMY"
};

const app = initializeApp(firebaseConfig);

export const db = getFirestore(app);
export const storage = getStorage(app);