import { auth } from "./firebase-config.js";

import {
    signInWithEmailAndPassword
} from "https://www.gstatic.com/firebasejs/12.0.0/firebase-auth.js";


const loginForm = document.getElementById("loginForm");

loginForm.addEventListener("submit", async (e) => {

    e.preventDefault();

    const email = document.getElementById("email").value.trim();

    const password = document.getElementById("password").value;

    try {

        // Login user
        await signInWithEmailAndPassword(auth, email, password);

        alert("Login Successful!");

        // Redirect to editor page
        window.location.href = "/main";

    }
    catch (error) {

        switch (error.code) {

            case "auth/invalid-credential":
                alert("Invalid email or password.");
                break;

            case "auth/user-not-found":
                alert("User not found.");
                break;

            case "auth/wrong-password":
                alert("Incorrect password.");
                break;

            case "auth/invalid-email":
                alert("Invalid email address.");
                break;

            case "auth/too-many-requests":
                alert("Too many failed attempts. Try again later.");
                break;

            default:
                alert(error.message);
        }

    }

});