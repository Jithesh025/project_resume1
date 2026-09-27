import { auth, db } from "./firebase-config.js";

import {
    createUserWithEmailAndPassword
} from "https://www.gstatic.com/firebasejs/12.0.0/firebase-auth.js";

import {
    doc,
    setDoc,
    serverTimestamp
} from "https://www.gstatic.com/firebasejs/12.0.0/firebase-firestore.js";


const registerForm = document.getElementById("registerForm");

registerForm.addEventListener("submit", async (e) => {

    e.preventDefault();

    const name = document.getElementById("name").value.trim();

    const email = document.getElementById("email").value.trim();

    const password = document.getElementById("password").value;

    const confirmPassword = document.getElementById("confirmPassword").value;

    // Validate passwords
    if (password !== confirmPassword) {
        alert("Passwords do not match.");
        return;
    }

    if (password.length < 6) {
        alert("Password must be at least 6 characters.");
        return;
    }

try {

    const userCredential = await createUserWithEmailAndPassword(
        auth,
        email,
        password
    );

    const user = userCredential.user;

    await setDoc(
        doc(db, "users", user.uid),
        {
            uid: user.uid,
            name: name,
            email: email,
            createdAt: serverTimestamp()
        }
    );

    alert("Registration Successful!");

    // Redirect to Flask login page
    window.location.href = "/";

}
catch (error) {

    console.error("Firebase Error:", error);

    switch (error.code) {

        case "auth/email-already-in-use":
            alert("Email already registered.");
            break;

        case "auth/invalid-email":
            alert("Invalid email address.");
            break;

        case "auth/weak-password":
            alert("Password must be at least 6 characters.");
            break;

        case "auth/network-request-failed":
            alert("Unable to connect to Firebase.");
            break;

        default:
            alert(error.message);
    }
}});