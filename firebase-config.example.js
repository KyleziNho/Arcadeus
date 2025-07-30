// Firebase Configuration Template
// 1. Create a new Firebase project at https://console.firebase.google.com
// 2. Enable Authentication and add Google and Microsoft as sign-in providers
// 3. Copy your Firebase config values here
// 4. Rename this file to firebase-config.js (don't commit the actual config!)

const firebaseConfig = {
    apiKey: "YOUR_API_KEY_HERE",
    authDomain: "YOUR_PROJECT_ID.firebaseapp.com",
    projectId: "YOUR_PROJECT_ID",
    storageBucket: "YOUR_PROJECT_ID.appspot.com",
    messagingSenderId: "YOUR_MESSAGING_SENDER_ID",
    appId: "YOUR_APP_ID"
};

// To enable Microsoft OAuth:
// 1. Go to Firebase Console > Authentication > Sign-in method
// 2. Enable Microsoft provider
// 3. Add your Microsoft App ID and App Secret
// 4. Copy the redirect URI and add it to your Microsoft app registration

// To enable Google OAuth:
// 1. Google OAuth is automatically configured with Firebase
// 2. Just enable it in Firebase Console > Authentication > Sign-in method

// For production, consider using environment variables or secure key management