// Firebase configuration - Update with your actual values from Firebase Console
const firebaseConfig = {
    apiKey: "YOUR_ACTUAL_API_KEY_FROM_FIREBASE", // Get this from Firebase Console > Project Settings > General > Web apps
    authDomain: "project-183306943540.firebaseapp.com", 
    projectId: "project-183306943540", 
    storageBucket: "project-183306943540.appspot.com",
    messagingSenderId: "YOUR_ACTUAL_MESSAGING_SENDER_ID", // Get this from Firebase Console
    appId: "YOUR_ACTUAL_APP_ID" // Get this from Firebase Console
};

// Initialize Firebase
firebase.initializeApp(firebaseConfig);
const auth = firebase.auth();

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
    setLoading(true);
    const provider = new firebase.auth.GoogleAuthProvider();
    
    try {
        const result = await auth.signInWithPopup(provider);
        console.log('Google sign in successful:', result.user.email);
    } catch (error) {
        console.error('Google sign in error:', error);
        showError('Failed to sign in with Google. Please try again.');
    } finally {
        setLoading(false);
    }
});

// Skip login functionality
document.getElementById('skipLogin')?.addEventListener('click', () => {
    // Create demo user data
    const demoUser = {
        uid: 'demo-user',
        email: 'demo@arcadeus.com',
        displayName: 'Demo User',
        photoURL: null,
        provider: 'demo'
    };
    
    // Store demo user info
    localStorage.setItem('arcadeusUser', JSON.stringify(demoUser));
    
    // Redirect to main app
    window.location.href = 'taskpane.html';
});

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
    if (!user && !window.location.href.includes('login.html')) {
        // Redirect to login if not authenticated
        window.location.href = 'login.html';
    }
    return user;
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
                <button class="profile-menu-item" id="viewProfile">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"></path>
                        <circle cx="12" cy="7" r="4"></circle>
                    </svg>
                    View Profile
                </button>
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
            
            viewProfile?.addEventListener('click', () => {
                // TODO: Implement profile view
                alert('Profile view coming soon!');
            });
            
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