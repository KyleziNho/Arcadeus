# Firebase Setup Guide for Arcadeus

## Step 1: Complete Firebase Console Setup

### A. Update Project Settings (from your screenshot):
1. **Public-facing name**: Change to `Arcadeus M&A Intelligence Suite`
2. **Support email**: Keep `osullivan7791@gmail.com`
3. Click **Save** to update these settings

### B. Get Your Firebase Configuration:
1. Go to Firebase Console → **Project Settings** (gear icon)
2. Scroll down to **"Your apps"** section
3. If you haven't added a web app yet:
   - Click **"</>"** (Web icon)
   - App nickname: `Arcadeus Web App`
   - Check **"Also set up Firebase Hosting"** (optional)
   - Click **Register app**

4. Copy the Firebase configuration object that looks like this:
```javascript
const firebaseConfig = {
  apiKey: "AIza...",
  authDomain: "project-183306943540.firebaseapp.com",
  projectId: "project-183306943540",
  storageBucket: "project-183306943540.appspot.com",
  messagingSenderId: "123456789",
  appId: "1:123456789:web:abc123"
};
```

## Step 2: Update Your Code

### Replace the config in `auth.js`:
```javascript
// Replace this section in auth.js with your actual values
const firebaseConfig = {
    apiKey: "YOUR_ACTUAL_API_KEY_HERE",
    authDomain: "project-183306943540.firebaseapp.com",
    projectId: "project-183306943540", 
    storageBucket: "project-183306943540.appspot.com",
    messagingSenderId: "YOUR_ACTUAL_SENDER_ID",
    appId: "YOUR_ACTUAL_APP_ID"
};
```

## Step 3: Test Authentication

1. Open your `login.html` in a browser
2. Click "Continue with Google"
3. You should see Google's sign-in popup
4. After signing in, you should be redirected to `taskpane.html`

## Troubleshooting

### If login doesn't work:

1. **Check Console Errors**: 
   - Open browser Developer Tools (F12)
   - Look for errors in Console tab

2. **Common Issues**:
   - **Invalid API Key**: Make sure you copied the full API key
   - **Domain not authorized**: Add your domain to Firebase Console → Authentication → Settings → Authorized domains
   - **App not found**: Make sure the projectId and appId are correct

3. **Test with localhost**:
   - Firebase automatically allows `localhost` for testing
   - Try opening `file:///path/to/your/login.html` first

### Enable Debug Mode:
Add this to your `auth.js` for debugging:
```javascript
// Add after firebase.initializeApp(firebaseConfig);
auth.onAuthStateChanged((user) => {
    console.log('Auth state changed:', user ? user.email : 'No user');
});
```

## Step 4: Production Setup

For production deployment:
1. Add your actual domain to **Authorized domains** in Firebase Console
2. Set up proper hosting (Firebase Hosting, Netlify, etc.)
3. Update any hardcoded URLs to use your production domain

## Security Notes

- Never commit your Firebase config to public repositories if it contains sensitive data
- The `apiKey` in Firebase web config is safe to expose (it's not a secret key)
- Real security comes from Firebase Security Rules, not hiding the config