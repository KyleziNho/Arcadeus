// Firebase configuration
const firebaseConfig = {
    apiKey: "AIzaSyA8btt-jKfkbf2_SspNZgz2mY9Fy6Jf-lo",
    authDomain: "arcadeus-5641b.firebaseapp.com", 
    projectId: "arcadeus-5641b", 
    storageBucket: "arcadeus-5641b.firebasestorage.app",
    messagingSenderId: "183306943540",
    appId: "1:183306943540:web:ed482fc5ccf6a4ecdede61",
    measurementId: "G-Z07WWE0H64"
};

// Initialize Firebase
firebase.initializeApp(firebaseConfig);
const auth = firebase.auth();
const db = firebase.firestore();

// Handle redirect result on page load
auth.getRedirectResult().then((result) => {
    if (result.credential) {
        // Google sign in successful via redirect
        console.log('Google sign in successful via redirect:', result.user.email);
    }
}).catch((error) => {
    console.error('Redirect sign in error:', error);
});

// Auth state observer
auth.onAuthStateChanged((user) => {
    if (user) {
        // User is signed in
        console.log('User signed in:', user.email);
        // Store user info in local storage for the add-in
        localStorage.setItem('arcadeusUser', JSON.stringify({
            uid: user.uid,
            email: user.email,
            displayName: user.displayName,
            photoURL: user.photoURL,
            provider: user.providerData[0].providerId
        }));
        // Redirect to main app or close window if in add-in context
        if (window.location.href.includes('login.html')) {
            window.location.href = 'taskpane.html';
        }
    } else {
        // User is signed out
        console.log('User signed out');
        localStorage.removeItem('arcadeusUser');
    }
});

// Show/hide loading overlay
function setLoading(isLoading) {
    const loadingOverlay = document.getElementById('loadingOverlay');
    if (loadingOverlay) {
        loadingOverlay.style.display = isLoading ? 'flex' : 'none';
    }
}

// Show error message
function showError(message) {
    const errorElement = document.getElementById('errorMessage');
    if (errorElement) {
        errorElement.textContent = message;
        errorElement.style.display = 'block';
        setTimeout(() => {
            errorElement.style.display = 'none';
        }, 5000);
    }
}

// Google Sign In
document.getElementById('googleSignIn')?.addEventListener('click', async () => {
    console.log('Google Sign In clicked');
    setLoading(true);
    
    // Check if Firebase is properly configured
    if (firebaseConfig.apiKey.includes('YOUR_ACTUAL')) {
        setLoading(false);
        showError('Firebase not configured properly.');
        return;
    }
    
    console.log('Firebase config looks valid, attempting sign in...');
    
    const provider = new firebase.auth.GoogleAuthProvider();
    
    try {
        // Try signInWithRedirect for better compatibility
        if (window.location.protocol === 'file:' || window.location.hostname === 'localhost') {
            // For local development, use redirect instead of popup
            await auth.signInWithRedirect(provider);
        } else {
            // For production, use popup
            const result = await auth.signInWithPopup(provider);
            console.log('Google sign in successful:', result.user.email);
        }
    } catch (error) {
        console.error('Google sign in error:', error);
        console.error('Error code:', error.code);
        console.error('Error message:', error.message);
        
        let errorMessage = 'Failed to sign in with Google. ';
        if (error.code === 'auth/api-key-not-valid.-please-pass-a-valid-api-key.') {
            errorMessage = 'API key issue detected. Please ensure:\n1. You\'ve enabled Google auth in Firebase Console\n2. The Web API Key is enabled in Google Cloud Console';
        } else if (error.code === 'auth/configuration-not-found' || error.code === 'auth/invalid-api-key') {
            errorMessage += 'Firebase configuration invalid.';
        } else if (error.code === 'auth/unauthorized-domain') {
            errorMessage = 'This domain is not authorized for OAuth operations. Please add your domain to Firebase Console → Authentication → Settings → Authorized domains.';
        } else if (error.code === 'auth/popup-blocked') {
            errorMessage += 'Popup was blocked. Trying redirect method...';
            // Fallback to redirect
            try {
                await auth.signInWithRedirect(provider);
                return;
            } catch (redirectError) {
                errorMessage = 'Authentication failed.';
            }
        } else if (error.code === 'auth/popup-closed-by-user') {
            errorMessage += 'Sign-in was cancelled.';
        } else {
            errorMessage += 'Error code: ' + error.code;
        }
        
        showError(errorMessage);
    } finally {
        setLoading(false);
    }
});

// Skip login functionality removed

// Microsoft authentication removed - using Google only

// Sign out function
function signOut() {
    auth.signOut().then(() => {
        console.log('Sign out successful');
        window.location.href = 'login.html';
    }).catch((error) => {
        console.error('Sign out error:', error);
    });
}

// Check authentication status
function checkAuth() {
    const user = JSON.parse(localStorage.getItem('arcadeusUser') || 'null');
    
    // Only redirect on taskpane.html, not on other pages
    if (!user && window.location.pathname.includes('taskpane.html')) {
        // Redirect to login if not authenticated
        window.location.href = 'login.html';
        return null;
    }
    
    return user;
}

// Check if user has completed onboarding
function hasCompletedOnboarding() {
    const onboardingStatus = localStorage.getItem('arcadeusOnboarding');
    return onboardingStatus === 'completed';
}

// Profile menu functions
function createProfileMenu(user) {
    if (!user) return '';
    
    return `
        <div class="profile-section">
            <button class="profile-button" id="profileButton">
                ${user.photoURL ? 
                    `<img src="${user.photoURL}" alt="${user.displayName}" class="profile-avatar">` :
                    `<div class="profile-avatar-placeholder">${(user.displayName || user.email)[0].toUpperCase()}</div>`
                }
            </button>
            <div class="profile-dropdown" id="profileDropdown">
                <div class="profile-info">
                    <div class="profile-name">${user.displayName || 'User'}</div>
                    <div class="profile-email">${user.email}</div>
                </div>
                <div class="profile-divider"></div>
                <a href="profile.html" class="profile-menu-item" style="text-decoration: none;">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"></path>
                        <circle cx="12" cy="7" r="4"></circle>
                    </svg>
                    View Profile
                </a>
                <button class="profile-menu-item" id="signOutButton">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"></path>
                        <polyline points="16 17 21 12 16 7"></polyline>
                        <line x1="21" y1="12" x2="9" y2="12"></line>
                    </svg>
                    Sign Out
                </button>
            </div>
        </div>
    `;
}

// Initialize profile menu
function initializeProfileMenu() {
    const user = checkAuth();
    if (user) {
        // Add profile menu to header if on main page
        const appHeader = document.querySelector('.app-header');
        if (appHeader) {
            const profileHTML = createProfileMenu(user);
            appHeader.insertAdjacentHTML('beforeend', profileHTML);
            
            // Add event listeners
            const profileButton = document.getElementById('profileButton');
            const profileDropdown = document.getElementById('profileDropdown');
            const signOutButton = document.getElementById('signOutButton');
            const viewProfile = document.getElementById('viewProfile');
            
            profileButton?.addEventListener('click', (e) => {
                e.stopPropagation();
                profileDropdown.classList.toggle('show');
            });
            
            signOutButton?.addEventListener('click', signOut);
            
            // Remove old viewProfile event listener since it's now a link
            
            // Close dropdown when clicking outside
            document.addEventListener('click', () => {
                profileDropdown?.classList.remove('show');
            });
        }
    }
}

// Initialize on page load
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initializeProfileMenu);
} else {
    initializeProfileMenu();
}

// Export functions for use in other scripts
window.arcadeusAuth = {
    checkAuth,
    signOut,
    getUser: () => JSON.parse(localStorage.getItem('arcadeusUser') || 'null')
};