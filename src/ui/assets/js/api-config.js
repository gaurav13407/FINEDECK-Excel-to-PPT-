/**
 * API Configuration for FinDeck Frontend
 * Centralizes all backend API endpoints and configuration
 */

class APIConfig {
    constructor() {
        // Backend server configuration
        this.BASE_URL = 'http://localhost:8000';
        this.API_VERSION = 'v1';
        this.API_BASE = `${this.BASE_URL}/api/${this.API_VERSION}`;
        
        // API endpoints
        this.endpoints = {
            // Authentication endpoints
            auth: {
                signup: `${this.API_BASE}/auth/signup`,
                login: `${this.API_BASE}/auth/login`,
                refresh: `${this.API_BASE}/auth/refresh`,
                logout: `${this.API_BASE}/auth/logout`,
                me: `${this.API_BASE}/auth/me`
            },
            
            // User management endpoints
            users: {
                // Backend exposes /users/me for profile get/update
                profile: `${this.API_BASE}/users/me`,
                updateProfile: `${this.API_BASE}/users/me`,
                subscription: `${this.API_BASE}/users/subscription`,
                // Use /users/stats for usage/statistics (backend exposes /stats)
                usage: `${this.API_BASE}/users/stats`
            },
            
            // File management endpoints
            files: {
                upload: `${this.API_BASE}/files/upload`,
                list: `${this.API_BASE}/files`,
                delete: (fileId) => `${this.API_BASE}/files/${fileId}`,
                download: (fileId) => `${this.API_BASE}/files/${fileId}/download`
            },
            
            // Conversion endpoints
            conversions: {
                convert: `${this.API_BASE}/conversions/convert`,
                templates: `${this.API_BASE}/conversions/templates`
            }
        };
        
        // Request headers
        this.defaultHeaders = {
            'Content-Type': 'application/json',
            'Accept': 'application/json'
        };
    }

    /**
     * Get authorization headers with JWT token
     */
    getAuthHeaders() {
        const token = localStorage.getItem('authToken');
        if (!token) {
            return this.defaultHeaders;
        }
        
        return {
            ...this.defaultHeaders,
            'Authorization': `Bearer ${token}`
        };
    }

    /**
     * Make API request with error handling
     */
    async makeRequest(url, options = {}) {
        try {
            const response = await fetch(url, {
                ...options,
                headers: {
                    ...this.getAuthHeaders(),
                    ...options.headers
                }
            });

            // Handle non-JSON responses (like file downloads)
            const contentType = response.headers.get('content-type');
            if (contentType && !contentType.includes('application/json')) {
                if (response.ok) {
                    return response;
                } else {
                    throw new Error(`HTTP ${response.status}: ${response.statusText}`);
                }
            }

            // Parse JSON response
            const data = await response.json();

            if (!response.ok) {
                throw new Error(data.detail || `HTTP ${response.status}: ${response.statusText}`);
            }

            return data;
        } catch (error) {
            console.error('API Request Error:', error);
            throw error;
        }
    }
}

// Create global API instance
console.log('📡 Loading API Config...');
window.apiConfig = new APIConfig();
console.log('✅ API Config loaded successfully:', window.apiConfig);