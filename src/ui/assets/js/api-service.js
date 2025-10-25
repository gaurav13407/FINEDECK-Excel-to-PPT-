/**
 * API Service for FinDeck Frontend
 * Handles all communication with the FastAPI backend
 */

class APIService {
    constructor() {
        // Minimal initialization; avoid noisy console output in production
        this.api = window.apiConfig || {};
    }

    // =================
    // AUTHENTICATION
    // =================

    /**
     * Register a new user
     */
    async signup(userData) {
        const payload = {
            name: userData.name,
            email: userData.email,
            password: userData.password
        };

        const response = await this.api.makeRequest(this.api.endpoints.auth.signup, {
            method: 'POST',
            body: JSON.stringify(payload)
        });

        // Store the user token if provided
        if (response.access_token) {
            localStorage.setItem('authToken', response.access_token);
        }

        return response;
    }

    /**
     * Login user
     */
    async login(email, password) {
        // FastAPI expects form data for OAuth2 login
        const formData = new FormData();
        formData.append('username', email); // Note: username field for email
        formData.append('password', password);

        const response = await fetch(this.api.endpoints.auth.login, {
            method: 'POST',
            body: formData // Don't set Content-Type, let browser set it for FormData
        });

        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.detail || 'Login failed');
        }

        const data = await response.json();
        
        // Store the access token
        if (data.access_token) {
            localStorage.setItem('authToken', data.access_token);
            localStorage.setItem('tokenType', data.token_type);
        }

        return data;
    }

    /**
     * Get current user profile
     */
    async getCurrentUser() {
        return await this.api.makeRequest(this.api.endpoints.auth.me);
    }

    /**
     * Logout user
     */
    async logout() {
        try {
            await this.api.makeRequest(this.api.endpoints.auth.logout, {
                method: 'POST'
            });
        } catch (error) {
            console.warn('Logout API call failed:', error);
        } finally {
            // Clear local storage regardless of API call result
            localStorage.removeItem('authToken');
            localStorage.removeItem('tokenType');
        }
    }

    // =================
    // USER MANAGEMENT
    // =================

    /**
     * Get user profile
     */
    async getUserProfile() {
        return await this.api.makeRequest(this.api.endpoints.users.profile);
    }

    /**
     * Update user profile
     */
    async updateUserProfile(profileData) {
        return await this.api.makeRequest(this.api.endpoints.users.updateProfile, {
            method: 'PUT',
            body: JSON.stringify(profileData)
        });
    }

    /**
     * Get user subscription details
     */
    async getUserSubscription() {
        return await this.api.makeRequest(this.api.endpoints.users.subscription);
    }

    /**
     * Update user subscription / plan
     * payload example: { plan: 'pro' }
     */
    async updateUserSubscription(payload) {
        return await this.api.makeRequest(this.api.endpoints.users.subscription, {
            method: 'PUT',
            body: JSON.stringify(payload)
        });
    }

    /**
     * Get user usage statistics
     */
    async getUserUsage() {
        return await this.api.makeRequest(this.api.endpoints.users.usage);
    }

    // =================
    // FILE MANAGEMENT
    // =================

    /**
     * Upload an Excel file
     */
    async uploadFile(file, metadata = {}) {
        const formData = new FormData();
        formData.append('file', file);
        
        // Add metadata fields
        Object.keys(metadata).forEach(key => {
            formData.append(key, metadata[key]);
        });

        const response = await fetch(this.api.endpoints.files.upload, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${localStorage.getItem('authToken')}`
            },
            body: formData
        });

        if (!response.ok) {
            const errorData = await response.json();
            console.log('❌ Upload failed:', {
                status: response.status,
                statusText: response.statusText,
                error: errorData,
                token: localStorage.getItem('authToken') ? 'Present' : 'Missing'
            });
            
            if (response.status === 401) {
                throw new Error('Authentication failed');
            }
            throw new Error(errorData.detail || 'File upload failed');
        }

        return await response.json();
    }

    /**
     * Get list of user files
     */
    async getFiles(skip = 0, limit = 20) {
        const url = `${this.api.endpoints.files.list}?skip=${skip}&limit=${limit}`;
        return await this.api.makeRequest(url);
    }

    /**
     * Delete a file
     */
    async deleteFile(fileId) {
        return await this.api.makeRequest(this.api.endpoints.files.delete(fileId), {
            method: 'DELETE'
        });
    }

    /**
     * Download a file
     */
    async downloadFile(fileId) {
        const response = await this.api.makeRequest(this.api.endpoints.files.download(fileId));
        return response; // This will be a blob for file downloads
    }

    // =================
    // CONVERSIONS
    // =================

    /**
     * Convert Excel file to PowerPoint
     */
    async convertExcelToPPT(fileId, conversionOptions = {}) {
        const formData = new FormData();
        formData.append('file_id', fileId);
        
        // Add conversion options
        if (conversionOptions.title) formData.append('title', conversionOptions.title);
        if (conversionOptions.subtitle) formData.append('subtitle', conversionOptions.subtitle);
        if (conversionOptions.sheet_name) formData.append('sheet_name', conversionOptions.sheet_name);
        if (conversionOptions.title_col) formData.append('title_col', conversionOptions.title_col);
        if (conversionOptions.mode) formData.append('mode', conversionOptions.mode);
        if (conversionOptions.limit) formData.append('limit', conversionOptions.limit.toString());

        const response = await fetch(this.api.endpoints.conversions.convert, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${localStorage.getItem('authToken')}`
            },
            body: formData
        });

        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.detail || 'Conversion failed');
        }

        // For file downloads, return the response directly
        return response;
    }

    /**
     * Get available PowerPoint templates
     */
    async getTemplates() {
        return await this.api.makeRequest(this.api.endpoints.conversions.templates);
    }

    // =================
    // UTILITY METHODS
    // =================

    /**
     * Check if user is authenticated
     */
    isAuthenticated() {
        return !!localStorage.getItem('authToken');
    }

    /**
     * Get stored auth token
     */
    getAuthToken() {
        return localStorage.getItem('authToken');
    }

    /**
     * Health check
     */
    async healthCheck() {
        try {
            const response = await fetch(`${this.api.BASE_URL}/health`);
            return response.ok;
        } catch (error) {
            return false;
        }
    }
}

// Create global API service instance
try {
    window.APIService = APIService; // Export the class
    window.apiService = new APIService(); // Create an instance
} catch (error) {
    // Keep only critical errors in the console
    console.error('Error creating APIService:', error);
}