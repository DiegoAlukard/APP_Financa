// Import the functions you need from the SDKs you need
import { initializeApp } from "firebase/app";
import { getAnalytics } from "firebase/analytics";
// TODO: Add SDKs for Firebase products that you want to use
// https://firebase.google.com/docs/web/setup#available-libraries

// Your web app's Firebase configuration
// For Firebase JS SDK v7.20.0 and later, measurementId is optional
const firebaseConfig = {
  apiKey: "AIzaSyDgrKwRpOZbULe_FjYbXSjCcHVOVa7disY",
  authDomain: "appfinanca-2e028.firebaseapp.com",
  projectId: "appfinanca-2e028",
  storageBucket: "appfinanca-2e028.firebasestorage.app",
  messagingSenderId: "734956292325",
  appId: "1:734956292325:web:738e071aedb0b8c312ac0c",
  measurementId: "G-Q8C43YLS8H"
};

// Initialize Firebase
const app = initializeApp(firebaseConfig);
const analytics = getAnalytics(app);
