// Simple login functionality without Firebase
console.log('🔧 Login page loaded - no external dependencies');

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

// Google Sign In - redirect to taskpane for Firebase handling
document.getElementById('googleSignIn')?.addEventListener('click', () => {
    console.log('Google Sign In clicked - redirecting to main app for Firebase auth');
    setLoading(true);
    
    // Set a flag to indicate Google auth was requested
    localStorage.setItem('arcadeusAuthIntent', 'google');
    
    // Redirect to main app where Firebase is available
    window.location.href = 'taskpane.html';
});

// Admin sign-in functionality - no Firebase needed
document.getElementById('adminSignIn')?.addEventListener('click', () => {
    console.log('Admin Sign In clicked');
    setLoading(true);
    
    const username = document.getElementById('adminUsername')?.value;
    const password = document.getElementById('adminPassword')?.value;
    
    console.log('Admin credentials entered:', { username, password: password ? '[HIDDEN]' : 'empty' });
    
    // Check admin credentials
    if (username === 'admin' && password === '88888888') {
        console.log('Admin credentials valid, proceeding with login...');
        
        // Create admin user object with pre-filled details
        const adminUser = {
            uid: 'admin-user-' + Date.now(),
            email: 'admin@arcadeus.com',
            displayName: 'Admin User',
            photoURL: null,
            provider: 'admin'
        };
        
        // Pre-filled onboarding data for admin
        const adminProfileData = {
            userType: 'company',
            organizationName: 'Arcadeus Development',
            userRole: 'System Administrator',
            teamSize: '21-50',
            propertyTypes: ['office', 'retail', 'multifamily', 'industrial'],
            onboardingCompleted: new Date().toISOString()
        };
        
        try {
            console.log('Storing admin data in localStorage...');
            
            // Store admin user info and profile data
            localStorage.setItem('arcadeusUser', JSON.stringify(adminUser));
            localStorage.setItem('arcadeusUserProfile', JSON.stringify(adminProfileData));
            localStorage.setItem('arcadeusOnboarding', 'completed');
            
            console.log('Admin user authenticated and stored:', adminUser.email);
            console.log('LocalStorage contents:', {
                user: localStorage.getItem('arcadeusUser'),
                profile: localStorage.getItem('arcadeusUserProfile'),
                onboarding: localStorage.getItem('arcadeusOnboarding')
            });
            
            // Redirect directly to main app (skip onboarding)
            console.log('Admin user redirecting to main app...');
            setLoading(false);
            window.location.href = 'taskpane.html';
            return;
            
        } catch (error) {
            console.error('Admin sign in error:', error);
            showError('Admin authentication failed: ' + error.message);
            setLoading(false);
            return;
        }
    } else {
        console.log('Invalid admin credentials provided');
        showError('Invalid admin credentials. Please check username and password.');
    }
    
    setLoading(false);
});