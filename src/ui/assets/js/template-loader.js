/**
 * Template Loader
 * Handles dynamic loading and display of templates from backend API
 */

class TemplateLoader {
    constructor() {
        this.apiBase = 'https://finedeck-excel-to-ppt-backend.onrender.com/api/v1';
        this.templates = [];
        this.selectedTemplate = localStorage.getItem('selectedTemplate') || null;
    }

    /**
     * Load all templates from backend API
     */
    async loadTemplates() {
        try {
            const token = localStorage.getItem('authToken');
            const headers = {};
            
            if (token) {
                headers['Authorization'] = `Bearer ${token}`;
            }
            
            const response = await fetch(`${this.apiBase}/templates`, {
                headers: headers
            });
            
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            
            this.templates = await response.json();
            return this.templates;
        } catch (error) {
            console.error('Error loading templates:', error);
            // Return hardcoded templates as fallback
            return this.getHardcodedTemplates();
        }
    }

    /**
     * Get hardcoded templates as fallback
     */
    getHardcodedTemplates() {
        return [
            {
                id: 'minimal_white',
                name: 'Minimal White',
                description: 'Clean minimalist design with black and blue accents, perfect for simple presentations',
                category: 'basic',
                source: 'built-in',
                image: 'assets/templates/Screenshot 2025-11-11 120233.png'
            },
            {
                id: 'corporate_blue',
                name: 'Corporate Blue',
                description: 'Classic corporate design with professional blue tones for business presentations',
                category: 'professional',
                source: 'built-in',
                image: 'assets/templates/Screenshot 2025-11-11 120323.png'
            },
            {
                id: 'modern_tech',
                name: 'Modern Tech',
                description: 'Sleek tech-focused design with vibrant colors for innovative presentations',
                category: 'professional',
                source: 'built-in',
                image: 'assets/templates/Screenshot 2025-11-11 120413.png'
            },
            {
                id: 'elegant_gray',
                name: 'Elegant Gray',
                description: 'Sophisticated gray palette with subtle red accents for professional settings',
                category: 'professional',
                source: 'built-in',
                image: 'assets/templates/Screenshot 2025-11-11 120505.png'
            },
            {
                id: 'ocean_blue',
                name: 'Ocean Blue',
                description: 'Calm ocean-inspired blues perfect for corporate and tech presentations',
                category: 'professional',
                source: 'built-in',
                image: 'assets/templates/Screenshot 2025-11-11 120542.png'
            },
            {
                id: 'dark_finance',
                name: 'Dark Finance',
                description: 'Professional dark theme with navy blue and gold accents, perfect for financial presentations',
                category: 'premium',
                source: 'built-in',
                image: 'assets/templates/Screenshot 2025-11-11 120603.png'
            },
            {
                id: 'vibrant_gradient',
                name: 'Vibrant Gradient',
                description: 'Bold design with purple and orange gradients for dynamic presentations',
                category: 'premium',
                source: 'built-in',
                image: 'assets/templates/Screenshot 2025-11-11 120615.png'
            },
            {
                id: 'sunset_orange',
                name: 'Sunset Orange',
                description: 'Warm orange and gold tones perfect for creative and dynamic presentations',
                category: 'premium',
                source: 'built-in',
                image: 'assets/templates/Screenshot 2025-11-11 120626.png'
            },
            {
                id: 'forest_green',
                name: 'Forest Green',
                description: 'Natural green tones ideal for sustainability and eco-friendly presentations',
                category: 'premium',
                source: 'built-in',
                image: 'assets/templates/Screenshot 2025-11-11 120639.png'
            },
            {
                id: 'royal_purple',
                name: 'Royal Purple',
                description: 'Luxurious deep purple theme perfect for creative and premium presentations',
                category: 'premium',
                source: 'built-in',
                image: 'assets/templates/Screenshot 2025-11-11 120648.png'
            }
        ];
    }

    /**
     * Get template by ID
     */
    async getTemplate(templateId) {
        try {
            const token = localStorage.getItem('authToken');
            const headers = {};
            
            if (token) {
                headers['Authorization'] = `Bearer ${token}`;
            }
            
            const response = await fetch(`${this.apiBase}/templates/${templateId}`, {
                headers: headers
            });
            
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            
            return await response.json();
        } catch (error) {
            console.error(`Error loading template ${templateId}:`, error);
            return null;
        }
    }

    /**
     * Select a template and save to localStorage
     */
    selectTemplate(templateId) {
        this.selectedTemplate = templateId;
        localStorage.setItem('selectedTemplate', templateId);
        console.log(`Template selected: ${templateId}`);
        
        // Dispatch custom event for other components to listen
        window.dispatchEvent(new CustomEvent('templateSelected', { 
            detail: { templateId: templateId }
        }));
    }

    /**
     * Get currently selected template
     */
    getSelectedTemplate() {
        return this.selectedTemplate;
    }

    /**
     * Upload custom template
     */
    async uploadTemplate(file) {
        try {
            const token = localStorage.getItem('authToken');
            if (!token) {
                throw new Error('Authentication required to upload templates');
            }
            
            const formData = new FormData();
            formData.append('file', file);
            
            const response = await fetch(`${this.apiBase}/templates/upload`, {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${token}`
                },
                body: formData
            });
            
            if (!response.ok) {
                const error = await response.json();
                throw new Error(error.detail || 'Upload failed');
            }
            
            return await response.json();
        } catch (error) {
            console.error('Error uploading template:', error);
            throw error;
        }
    }
}

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
    module.exports = TemplateLoader;
}
