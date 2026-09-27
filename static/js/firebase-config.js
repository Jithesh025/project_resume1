// ==========================================
// Firebase SDK
// ==========================================

import {
    initializeApp
} from "https://www.gstatic.com/firebasejs/12.0.0/firebase-app.js";

import {
    getAuth
} from "https://www.gstatic.com/firebasejs/12.0.0/firebase-auth.js";

import {
    getFirestore
} from "https://www.gstatic.com/firebasejs/12.0.0/firebase-firestore.js";


// ==========================================
// Firebase Configuration
// ==========================================

const firebaseConfig = {

    apiKey: "AIzaSyBNPIO0xSLWEB1yssY2AaTJ43e4hJjw0R4",

    authDomain: "airesumefy.firebaseapp.com",

    projectId: "airesumefy",

    storageBucket: "airesumefy.firebasestorage.app",

    messagingSenderId: "463166080179",

    appId: "1:463166080179:web:e0565e204d9d4d7bf91cc6",

    measurementId: "G-R992PXKKS5"
};


// ==========================================
// Initialize Firebase
// ==========================================

const app = initializeApp(firebaseConfig);


// ==========================================
// Firebase Authentication
// ==========================================

const auth = getAuth(app);


// ==========================================
// Firestore Database
// ==========================================

const db = getFirestore(app);


// ==========================================
// Export
// ==========================================

export {
    auth,
    db
};