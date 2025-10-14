/**
 * FinDeck Authentication Manager
 * Handles user authentication state across all pages
 */

class AuthManager {
    constructor() {
        this.isLoggedIn = false;
        this.currentUser = null;
        this.init();
    }

    init() {
        this.checkAuthState();
        this.updateUI();
        
        // Also listen for page visibility changes to refresh auth state
        document.addEventListener('visibilitychange', () => {
            if (!document.hidden) {
                this.checkAuthState();
                this.updateUI();
            }
        });
        
        // Listen for focus events to refresh auth state when user returns to tab
        window.addEventListener('focus', () => {
            this.checkAuthState();
            this.updateUI();
        });
    }

    // Check if user is logged in by checking localStorage
    checkAuthState() {
        const authData = localStorage.getItem('finDeckAuth');
        if (authData) {
            try {
                const parsedData = JSON.parse(authData);
                // Check if session is still valid (24 hours)
                const sessionTime = new Date(parsedData.loginTime);
                const now = new Date();
                const timeDiff = now - sessionTime;
                const hoursDiff = timeDiff / (1000 * 60 * 60);

                if (hoursDiff < 24) {
                    this.isLoggedIn = true;
                    this.currentUser = parsedData.user;
                } else {
                    // Session expired
                    this.logout();
                }
            } catch (error) {
                console.error('Error parsing auth data:', error);
                this.logout();
            }
        }
    }

    // Login user and store session
    login(userData) {
        const authData = {
            user: userData,
            loginTime: new Date().toISOString(),
            sessionId: this.generateSessionId()
        };

        localStorage.setItem('finDeckAuth', JSON.stringify(authData));
        this.isLoggedIn = true;
        this.currentUser = userData;
        this.updateUI();

        // Dispatch custom event for other components
        window.dispatchEvent(new CustomEvent('authStateChanged', {
            detail: { isLoggedIn: true, user: userData }
        }));

        return true;
    }

    // Logout user and clear session
    logout() {
        localStorage.removeItem('finDeckAuth');
        this.isLoggedIn = false;
        this.currentUser = null;
        this.updateUI();

        // Dispatch custom event
        window.dispatchEvent(new CustomEvent('authStateChanged', {
            detail: { isLoggedIn: false, user: null }
        }));

        // Redirect to home page if on protected pages
        const protectedPages = ['mainpage.html', 'dashboard.html', 'templates.html'];
        const currentPage = window.location.pathname.split('/').pop();
        if (protectedPages.includes(currentPage)) {
            window.location.href = 'index.html';
        }
    }

    // Update UI based on authentication state
    updateUI() {
        this.updateNavigation();
        this.updateHeroSection();
        this.showUserProfile();
    }

    // Update navigation for logged in/out state
    updateNavigation() {
        const loginBtn = document.querySelector('.btn-getstarted.btn-login');
        const userMenuContainer = document.querySelector('.user-menu-container');

        if (this.isLoggedIn && this.currentUser) {
            // Hide login button
            if (loginBtn) {
                loginBtn.style.display = 'none';
            }

            // Show user menu or create it
            this.createUserMenu();
        } else {
            // Show login button
            if (loginBtn) {
                loginBtn.style.display = 'inline-block';
                loginBtn.textContent = 'Login/Sign Up';
                loginBtn.href = 'login.html';
            }

            // Hide user menu
            if (userMenuContainer) {
                userMenuContainer.remove();
            }
        }
    }

    // Force refresh authentication state and UI
    forceRefresh() {
        this.checkAuthState();
        this.updateUI();
    }

    // Create user menu dropdown for logged in users
    createUserMenu() {
        const header = document.querySelector('#header .container');
        if (!header) return;

        // Remove existing user menu
        const existingMenu = document.querySelector('.user-menu-container');
        if (existingMenu) {
            existingMenu.remove();
        }

        const userMenuHTML = `
            <div class="user-menu-container">
                <div class="user-dropdown">
                    <button class="user-btn" id="userMenuBtn">
                        <img src="https://via.placeholder.com/32x32/e59d02/ffffff?text=${this.getUserInitials()}" 
                             alt="Profile" class="user-avatar">
                        <span class="user-name">${this.currentUser.name || 'User'}</span>
                        <i class="bi bi-chevron-down dropdown-arrow"></i>
                    </button>
                    
                    <div class="user-dropdown-menu" id="userDropdownMenu">
                        <div class="user-info">
                            <img src="https://via.placeholder.com/48x48/e59d02/ffffff?text=${this.getUserInitials()}" 
                                 alt="Profile" class="user-avatar-large">
                            <div class="user-details">
                                <h4>${this.currentUser.name || 'User'}</h4>
                                <p>${this.currentUser.email || ''}</p>
                                <span class="user-plan">${this.currentUser.plan || 'Free'} Plan</span>
                            </div>
                        </div>
                        
                        <hr class="dropdown-divider">
                        
                        <a href="mainpage.html" class="dropdown-item">
                            <i class="bi bi-house"></i>
                            Dashboard
                        </a>
                        
                        <a href="templates.html" class="dropdown-item">
                            <i class="bi bi-collection"></i>
                            Templates
                        </a>
                        
                        <a href="help.html" class="dropdown-item">
                            <i class="bi bi-question-circle"></i>
                            Help & Support
                        </a>
                        
                        <hr class="dropdown-divider">
                        
                        <a href="#" class="dropdown-item" onclick="authManager.logout()">
                            <i class="bi bi-box-arrow-right"></i>
                            Sign Out
                        </a>
                    </div>
                </div>
            </div>
        `;

        header.insertAdjacentHTML('beforeend', userMenuHTML);
        this.initializeUserMenu();
    }

    // Initialize user menu interactions
    initializeUserMenu() {
        const userBtn = document.getElementById('userMenuBtn');
        const userDropdown = document.getElementById('userDropdownMenu');

        if (!userBtn || !userDropdown) return;

        userBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            userDropdown.classList.toggle('show');
        });

        // Close dropdown when clicking outside
        document.addEventListener('click', (e) => {
            if (!userBtn.contains(e.target) && !userDropdown.contains(e.target)) {
                userDropdown.classList.remove('show');
            }
        });

        // Prevent dropdown from closing when clicking inside
        userDropdown.addEventListener('click', (e) => {
            e.stopPropagation();
        });
    }

    // Update hero section for logged in users
    updateHeroSection() {
        const heroSection = document.querySelector('#hero');
        const heroTitle = document.querySelector('.hero-title');
        const heroDescription = document.querySelector('.hero-description');
        const ctaButtons = document.querySelector('.cta-buttons');

        if (this.isLoggedIn && heroSection) {
            // Update hero content for logged in users
            if (heroTitle) {
                heroTitle.innerHTML = `Welcome back, ${this.currentUser.name || 'User'}!<br>
                    <span class="typed" data-typed-items="Ready to Convert?, Create Presentations, Upload Files, Start Converting"></span>`;
            }

            if (heroDescription) {
                heroDescription.textContent = "Continue where you left off or start a new Excel to PowerPoint conversion project.";
            }

            // Update CTA buttons
            this.updateHeroCTAButtons();
        }
    }

    // Update hero CTA buttons for logged in users
    updateHeroCTAButtons() {
        const ctaButtons = document.querySelector('.cta-buttons');
        if (!ctaButtons) return;

        ctaButtons.innerHTML = `
            <a href="mainpage.html" class="btn btn-primary btn-lg me-3" data-aos="fade-up" data-aos-delay="400">
                <i class="bi bi-upload me-2"></i>
                Start Converting
            </a>
            <a href="dashboard.html" class="btn btn-outline-primary btn-lg" data-aos="fade-up" data-aos-delay="500">
                <i class="bi bi-speedometer2 me-2"></i>
                View Dashboard
            </a>
        `;
    }

    // Show user profile section if logged in
    showUserProfile() {
        const profileSection = document.querySelector('#user-profile');
        
        if (this.isLoggedIn && profileSection) {
            profileSection.style.display = 'block';
            this.updateProfileSection();
        } else if (profileSection) {
            profileSection.style.display = 'none';
        }
    }

    // Update profile section content
    updateProfileSection() {
        const profileSection = document.querySelector('#user-profile');
        if (!profileSection || !this.currentUser) return;

        profileSection.innerHTML = `
            <div class="container">
                <div class="row">
                    <div class="col-lg-8 mx-auto text-center">
                        <h2>Your Account</h2>
                        <div class="profile-info mt-4">
                            <div class="row">
                                <div class="col-md-4">
                                    <div class="stat-card">
                                        <h3>${this.currentUser.conversions || 0}</h3>
                                        <p>Total Conversions</p>
                                    </div>
                                </div>
                                <div class="col-md-4">
                                    <div class="stat-card">
                                        <h3>${this.currentUser.templates || 0}</h3>
                                        <p>Templates Used</p>
                                    </div>
                                </div>
                                <div class="col-md-4">
                                    <div class="stat-card">
                                        <h3>${this.currentUser.plan || 'Free'}</h3>
                                        <p>Current Plan</p>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        `;
    }

    // Get user initials for avatar
    getUserInitials() {
        if (!this.currentUser || !this.currentUser.name) return 'U';
        
        const names = this.currentUser.name.split(' ');
        if (names.length >= 2) {
            return names[0].charAt(0) + names[1].charAt(0);
        }
        return names[0].charAt(0);
    }

    // Generate session ID
    generateSessionId() {
        return Math.random().toString(36).substring(2) + Date.now().toString(36);
    }

    // Check if user is on a protected page
    isProtectedPage() {
        const protectedPages = ['mainpage.html', 'dashboard.html', 'templates.html'];
        const currentPage = window.location.pathname.split('/').pop();
        return protectedPages.includes(currentPage);
    }

    // Redirect to login if not authenticated on protected pages
    requireAuth() {
        if (!this.isLoggedIn && this.isProtectedPage()) {
            window.location.href = 'login.html?redirect=' + encodeURIComponent(window.location.pathname);
            return false;
        }
        return true;
    }

    // Get current user data
    getCurrentUser() {
        return this.currentUser;
    }

    // Update user data
    updateUser(userData) {
        if (!this.isLoggedIn) return false;

        this.currentUser = { ...this.currentUser, ...userData };
        
        const authData = JSON.parse(localStorage.getItem('finDeckAuth'));
        authData.user = this.currentUser;
        localStorage.setItem('finDeckAuth', JSON.stringify(authData));
        
        this.updateUI();
        return true;
    }
}

// Initialize auth manager globally
window.authManager = new AuthManager();

// Export for modules
if (typeof module !== 'undefined' && module.exports) {
    module.exports = AuthManager;
}