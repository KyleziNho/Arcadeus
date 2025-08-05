// Simple login functionality without Firebase
console.log('🔧 Login page loaded - no external dependencies');

// Show any error messages from previous attempts
window.addEventListener('DOMContentLoaded', () => {
    const errorMessage = localStorage.getItem('arcadeusLoginError');
    if (errorMessage) {
        showError(errorMessage);
        localStorage.removeItem('arcadeusLoginError'); // Clear the error
        
        // Make admin login more prominent when Google fails
        const adminToggle = document.getElementById('adminLoginToggle');
        const googleButton = document.getElementById('googleSignIn');
        
        if (adminToggle) {
            adminToggle.style.fontSize = '14px';
            adminToggle.style.fontWeight = '600';
            adminToggle.style.color = '#667eea';
            adminToggle.style.textDecoration = 'none';
            adminToggle.style.padding = '8px 16px';
            adminToggle.style.border = '2px solid #667eea';
            adminToggle.style.borderRadius = '6px';
            adminToggle.style.background = 'rgba(102, 126, 234, 0.05)';
            adminToggle.textContent = '👨‍💼 Admin Login (Recommended for Excel Online)';
            
            // Auto-show the admin form
            setTimeout(() => {
                const adminForm = document.getElementById('adminLoginForm');
                if (adminForm) {
                    adminForm.style.display = 'block';
                }
            }, 1000);
        }
        
        if (googleButton) {
            googleButton.style.opacity = '0.6';
            googleButton.style.cursor = 'not-allowed';
            googleButton.title = 'Not available in Excel Online - Use Admin Login below';
        }
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

// Google Sign In is now handled directly in auth.js
// No need for redirect logic

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