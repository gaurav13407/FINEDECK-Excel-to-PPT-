/**
 * FinDeck Main Page JavaScript
 * Handles file upload, profile management, navigation, and UI interactions
 */

class FinDeckApp {
    constructor() {
        this.uploadedFiles = [];
        this.currentStep = 1;
        this.maxFileSize = 50 * 1024 * 1024; // 50MB
        this.allowedTypes = [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
            'application/vnd.ms-excel', // .xls
            'text/csv' // .csv
        ];
        
        // Initialize API service with retry mechanism
        this.initializeApiService();
        
        this.init();
    }

    initializeApiService() {
        // Check if APIService is available
        if (window.APIService) {
            this.apiService = new window.APIService();
            console.log('✅ API Service initialized successfully');
        } else {
            console.warn('⚠️ APIService not available, retrying in 500ms...');
            // Retry after a short delay to allow scripts to load
            setTimeout(() => {
                if (window.APIService) {
                    this.apiService = new window.APIService();
                    console.log('✅ API Service initialized successfully (retry)');
                } else {
                    console.error('❌ APIService still not available');
                    this.apiService = null;
                }
            }, 500);
        }
    }

    init() {
        this.initializeEventListeners();
        this.initializeFileUpload();
        this.initializeProfileDropdown();
        this.initializeModals();
        this.initializeStepIndicator();
        this.initializeProgressTracking();
        // Load billing/subscription info
        this.subscriptionData = null;
        this.loadSubscriptionInfo();
        // Wire update subscription button
        this.initSubscriptionButtons();
    }

    initializeEventListeners() {
        document.addEventListener('DOMContentLoaded', () => {
            console.log('FinDeck Application Initialized');
            this.updateUIState();
        });

        // Handle keyboard shortcuts
        document.addEventListener('keydown', (e) => {
            if (e.ctrlKey && e.key === 'u') {
                e.preventDefault();
                this.openFileDialog();
            }
            if (e.key === 'Escape') {
                this.closeAllModals();
                this.closeProfileDropdown();
            }
        });

        // Handle window resize
        window.addEventListener('resize', () => {
            this.handleResize();
        });
    }

    initializeConvertButton() {
        // Setup convert button - call this after DOM is ready
        const convertBtn = document.getElementById('convertBtn');
        if (convertBtn) {
            console.log('✅ Convert button found, setting up event listener');
            convertBtn.addEventListener('click', (e) => {
                e.preventDefault();
                console.log('🔄 Convert button clicked');
                this.startConversion();
            });
        } else {
            console.log('❌ Convert button not found');
        }
    }

    initializeFileUpload() {
        const uploadArea = document.getElementById('uploadArea');
        const fileInput = document.getElementById('fileInput');
        const browseBtn = document.querySelector('.browse-btn');

        if (!uploadArea || !fileInput) return;

        // Drag and drop functionality
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            e.stopPropagation();
            uploadArea.classList.add('drag-over');
        });

        uploadArea.addEventListener('dragenter', (e) => {
            e.preventDefault();
            e.stopPropagation();
            uploadArea.classList.add('drag-over');
        });

        uploadArea.addEventListener('dragleave', (e) => {
            e.preventDefault();
            e.stopPropagation();
            // Only remove if leaving the upload area completely
            if (!uploadArea.contains(e.relatedTarget)) {
                uploadArea.classList.remove('drag-over');
            }
        });

        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            e.stopPropagation();
            uploadArea.classList.remove('drag-over');
            
            const files = Array.from(e.dataTransfer.files);
            this.handleFiles(files);
        });

        // Click to upload (only on upload area, not on browse button)
        uploadArea.addEventListener('click', (e) => {
            // Don't trigger if clicking on browse button or its children
            if (e.target.closest('.browse-btn')) {
                return; // Let the browse button handle it
            }
            console.log('🔄 Upload area clicked, opening file dialog');
            this.openFileDialog();
        });

        // Browse button
        if (browseBtn) {
            browseBtn.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                console.log('🔄 Browse button clicked, opening file dialog');
                this.openFileDialog();
            });
        }

        // File input change
        fileInput.addEventListener('change', (e) => {
            console.log('🔄 File input changed, files selected:', e.target.files.length);
            const files = Array.from(e.target.files);
            console.log('📁 Selected files:', files.map(f => `${f.name} (${f.type}, ${f.size} bytes)`));
            this.handleFiles(files);
        });
    }

    openFileDialog() {
        console.log('🔄 Opening file dialog...');
        const fileInput = document.getElementById('fileInput');
        if (fileInput) {
            console.log('✅ File input found, triggering click');
            fileInput.click();
        } else {
            console.error('❌ File input element not found!');
        }
    }

    async handleFiles(files) {
        console.log('🔄 handleFiles called with:', files);
        
        if (!files || files.length === 0) {
            console.log('❌ No files provided');
            return;
        }

        const validFiles = [];
        const errors = [];

        files.forEach(file => {
            console.log('🔍 Validating file:', file.name, file.type, file.size);
            const validation = this.validateFile(file);
            if (validation.valid) {
                validFiles.push(file);
                console.log('✅ File valid:', file.name);
            } else {
                errors.push({
                    fileName: file.name,
                    error: validation.error
                });
                console.log('❌ File invalid:', file.name, validation.error);
            }
        });

        if (errors.length > 0) {
            console.log('❌ File validation errors:', errors);
            this.showFileErrors(errors);
        }

        if (validFiles.length > 0) {
            console.log('📤 Starting upload for valid files:', validFiles.map(f => f.name));
            // Upload files to backend
            await this.uploadFilesToBackend(validFiles);
        } else {
            console.log('❌ No valid files to upload');
        }
    }

    async uploadFilesToBackend(files) {
        // Check API service availability
        if (!this.apiService) {
            console.error('❌ API service not available');
            this.showNotification('Error', 'API service not available. Please refresh the page and try again.', 'error');
            return;
        }

        // Check authentication
        const token = localStorage.getItem('authToken');
        if (!token) {
            console.error('❌ No authentication token found');
            this.showNotification('Error', 'Please log in again to upload files.', 'error');
            return;
        }

        console.log('🔄 Starting file upload...', files);
        this.showUploadProgress();

        try {
            for (let i = 0; i < files.length; i++) {
                const file = files[i];
                console.log(`📤 Uploading file ${i + 1}/${files.length}: ${file.name}`);
                this.updateUploadProgress(i + 1, files.length, `Uploading ${file.name}...`);
                
                const response = await this.apiService.uploadFile(file);
                console.log('✅ Upload response:', response);
                
                if (response) {
                    // Add file metadata from backend response
                    const fileWithId = {
                        ...file,
                        id: response._id || response.id || `temp_${Date.now()}_${i}`,
                        _id: response._id || response.id,
                        url: response.url || null,
                        uploaded: true,
                        // Store the backend response for conversion
                        backendResponse: response
                    };
                    this.uploadedFiles.push(fileWithId);
                    console.log('📁 File added to uploaded list:', fileWithId);
                } else {
                    throw new Error(`No response received for ${file.name}`);
                }
            }

            console.log('🎉 All files uploaded successfully');
            this.hideUploadProgress();
            this.displayUploadedFiles();
            this.updateStepIndicator(2);
            this.showFileListSection();
            this.showNotification('Success', `Successfully uploaded ${files.length} file(s)!`, 'success');

        } catch (error) {
            console.error('❌ Upload error:', error);
            this.hideUploadProgress();
            this.showNotification('Error', error.message || 'Failed to upload files. Please try again.', 'error');
        }
    }

    showUploadProgress() {
        const uploadSection = document.getElementById('uploadSection');
        if (uploadSection) {
            uploadSection.innerHTML = `
                <div class="upload-progress-container">
                    <h3>Uploading Files...</h3>
                    <div class="progress-container">
                        <div class="progress-bar" id="uploadProgress"></div>
                    </div>
                    <p class="progress-text" id="uploadProgressText">Preparing upload...</p>
                </div>
            `;
        }
    }

    updateUploadProgress(current, total, message) {
        const progressBar = document.getElementById('uploadProgress');
        const progressText = document.getElementById('uploadProgressText');
        
        const percentage = (current / total) * 100;
        
        if (progressBar) {
            progressBar.style.width = `${percentage}%`;
        }
        
        if (progressText) {
            progressText.textContent = message || `Uploading ${current} of ${total} files...`;
        }
    }

    hideUploadProgress() {
        // Progress will be hidden when we transition to file list section
    }

    validateFile(file) {
        // Check file type
        if (!this.allowedTypes.includes(file.type)) {
            return {
                valid: false,
                error: 'Invalid file type. Please upload Excel (.xlsx, .xls) or CSV files.'
            };
        }

        // Check file size
        if (file.size > this.maxFileSize) {
            return {
                valid: false,
                error: 'File size exceeds 50MB limit.'
            };
        }

        // Check if file already exists
        if (this.uploadedFiles.some(f => f.name === file.name && f.size === file.size)) {
            return {
                valid: false,
                error: 'File already uploaded.'
            };
        }

        return { valid: true };
    }

    showFileErrors(errors) {
        const errorMessages = errors.map(err => `${err.fileName}: ${err.error}`).join('\n');
        this.showNotification('Error', errorMessages, 'error');
    }

    displayUploadedFiles() {
        const fileList = document.getElementById('uploadedFilesList');
        if (!fileList) return;

        fileList.innerHTML = '';

        this.uploadedFiles.forEach((file, index) => {
            const fileItem = this.createFileItem(file, index);
            fileList.appendChild(fileItem);
        });

        this.updateFileStats();
    }

    createFileItem(file, index) {
        const fileItem = document.createElement('div');
        fileItem.className = 'file-item';
        fileItem.innerHTML = `
            <div class="file-info">
                <div class="file-icon">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                        <polyline points="14,2 14,8 20,8"/>
                        <line x1="16" y1="13" x2="8" y2="13"/>
                        <line x1="16" y1="17" x2="8" y2="17"/>
                        <polyline points="10,9 9,9 8,9"/>
                    </svg>
                </div>
                <div class="file-details">
                    <h4>${file.name}</h4>
                    <p>${this.formatFileSize(file.size)} • ${this.getFileType(file.type)}</p>
                </div>
            </div>
            <div class="file-actions">
                <button class="action-btn preview-btn" onclick="finDeckApp.previewFile(${index})" title="Preview">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/>
                        <circle cx="12" cy="12" r="3"/>
                    </svg>
                </button>
                <button class="action-btn remove-btn" onclick="finDeckApp.removeFile(${index})" title="Remove">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <polyline points="3,6 5,6 21,6"/>
                        <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2 2h4a2 2 0 0 1 2 2v2"/>
                        <line x1="10" y1="11" x2="10" y2="17"/>
                        <line x1="14" y1="11" x2="14" y2="17"/>
                    </svg>
                </button>
            </div>
        `;

        return fileItem;
    }

    removeFile(index) {
        if (index >= 0 && index < this.uploadedFiles.length) {
            this.uploadedFiles.splice(index, 1);
            this.displayUploadedFiles();
            
            if (this.uploadedFiles.length === 0) {
                this.showUploadSection();
                this.updateStepIndicator(1);
            }
        }
    }

    previewFile(index) {
        const file = this.uploadedFiles[index];
        if (!file) return;

        // In a real application, you would show a preview modal
        this.showNotification('Preview', `Previewing ${file.name}...`, 'info');
        
        // Simulate file preview
        setTimeout(() => {
            this.showNotification('Success', 'File preview loaded successfully!', 'success');
        }, 1000);
    }

    formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    getFileType(mimeType) {
        const types = {
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'Excel (.xlsx)',
            'application/vnd.ms-excel': 'Excel (.xls)',
            'text/csv': 'CSV'
        };
        return types[mimeType] || 'Unknown';
    }

    updateFileStats() {
        const totalFiles = this.uploadedFiles.length;
        const totalSize = this.uploadedFiles.reduce((sum, file) => sum + file.size, 0);
        
        const statsElement = document.getElementById('fileStats');
        if (statsElement) {
            statsElement.innerHTML = `
                <span>${totalFiles} file${totalFiles !== 1 ? 's' : ''}</span>
                <span>•</span>
                <span>${this.formatFileSize(totalSize)}</span>
            `;
        }
    }

    showFileListSection() {
        const fileListSection = document.getElementById('fileListSection');
        const uploadSection = document.getElementById('uploadSection');
        
        if (fileListSection) {
            fileListSection.style.display = 'block';
            fileListSection.classList.add('fade-in');
        }
        if (uploadSection) {
            uploadSection.style.display = 'none';
        }
    }

    showUploadSection() {
        const fileListSection = document.getElementById('fileListSection');
        const uploadSection = document.getElementById('uploadSection');
        
        if (uploadSection) {
            uploadSection.style.display = 'block';
            uploadSection.classList.add('fade-in');
        }
        if (fileListSection) {
            fileListSection.style.display = 'none';
        }
    }

    initializeStepIndicator() {
        const steps = document.querySelectorAll('.step');
        steps.forEach((step, index) => {
            step.addEventListener('click', () => {
                if (index + 1 <= this.currentStep) {
                    this.updateStepIndicator(index + 1);
                }
            });
        });
    }

    updateStepIndicator(step) {
        this.currentStep = step;
        const steps = document.querySelectorAll('.step');
        
        steps.forEach((stepElement, index) => {
            stepElement.classList.remove('active', 'completed');
            
            if (index + 1 === step) {
                stepElement.classList.add('active');
            } else if (index + 1 < step) {
                stepElement.classList.add('completed');
            }
        });

        this.updateStepContent(step);
    }

    updateStepContent(step) {
        const sections = {
            1: 'uploadSection',
            2: 'fileListSection',
            3: 'settingsSection',
            4: 'resultsSection'
        };

        Object.values(sections).forEach(sectionId => {
            const section = document.getElementById(sectionId);
            if (section) {
                section.style.display = 'none';
            }
        });

        const activeSection = document.getElementById(sections[step]);
        if (activeSection) {
            activeSection.style.display = 'block';
            activeSection.classList.add('fade-in');
        }
    }

    initializeProfileDropdown() {
        const profileBtn = document.getElementById('profileBtn');
        const profileDropdown = document.getElementById('profileDropdown');

        if (!profileBtn || !profileDropdown) return;

        profileBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            this.toggleProfileDropdown();
        });

        // Close dropdown when clicking outside
        document.addEventListener('click', (event) => {
            if (!profileBtn.contains(event.target) && !profileDropdown.contains(event.target)) {
                this.closeProfileDropdown();
            }
        });

        // Prevent dropdown from closing when clicking inside
        profileDropdown.addEventListener('click', (e) => {
            e.stopPropagation();
        });
        
        // Initialize user profile
        this.initializeUserProfile();
    }

    initializeUserProfile() {
        // Get user info from localStorage or API
        const userInfo = this.getUserInfo();
        this.updateProfileDisplay(userInfo);
        // Listen for auth changes so profile and plan update immediately
        window.addEventListener('authStateChanged', (e) => {
            try {
                const newUser = e && e.detail && e.detail.user ? e.detail.user : this.getUserInfo();
                this.updateProfileDisplay(newUser);
                // refresh subscription info when auth changes
                this.loadSubscriptionInfo();
            } catch (err) { console.warn('authStateChanged handler error', err); }
        });
    }

    getUserInfo() {
        // Prefer the global auth manager if available
        try {
            if (window.authManager && typeof window.authManager.getCurrentUser === 'function') {
                const u = window.authManager.getCurrentUser();
                if (u) return u;
            }
        } catch (err) { /* ignore */ }

        // Next, try finDeckAuth stored session structure
        try {
            const authData = localStorage.getItem('finDeckAuth');
            if (authData) {
                const parsed = JSON.parse(authData);
                if (parsed && parsed.user) return parsed.user;
            }
        } catch (err) { /* ignore */ }

        // Fallback to legacy/currentUser key
        try {
            const storedUser = localStorage.getItem('currentUser');
            if (storedUser) return JSON.parse(storedUser);
        } catch (e) {
            console.error('Error parsing stored user info:', e);
        }

        // Default user info
        return {
            name: 'Guest User',
            email: '',
            plan: 'AI Pro',
            avatar: null // No custom avatar, will use default
        };
    }

    updateProfileDisplay(userInfo) {
        // Update profile name
        const profileName = document.getElementById('profileName');
        const userName = document.getElementById('userName');
        const userEmail = document.getElementById('userEmail');
        const userPlan = document.getElementById('userPlan');

        if (profileName) profileName.textContent = userInfo.name;
        if (userName) userName.textContent = userInfo.name;
        if (userEmail) userEmail.textContent = userInfo.email;
        if (userPlan) {
            const p = userInfo.plan || '';
            userPlan.textContent = p + (p.toLowerCase().includes('plan') ? '' : ' Plan');
        }

        // Update mainpage-specific plan labels only (avoid global catch-alls)
        try {
            const mainHeaderPlan = document.getElementById('userPlan');
            const mainDropdownPlan = document.getElementById('dropdownUserPlan');
            const text = (userInfo.plan || '');
            const out = text + (text.toLowerCase().includes('plan') ? '' : ' Plan');
            if (mainHeaderPlan) mainHeaderPlan.textContent = out;
            if (mainDropdownPlan) mainDropdownPlan.textContent = out;
        } catch (err) { /* ignore DOM issues */ }

        // Update avatars (default SVG avatars are already in HTML)
        // If user has custom avatar, we could update it here
        if (userInfo.avatar) {
            this.updateCustomAvatar(userInfo.avatar);
        }
        // Default SVG avatars are already set in HTML and styled with CSS
    }


    // Simple auth check used by plan CTAs
    isAuthenticated() {
        try {
            if (window.authManager && typeof window.authManager.isAuthenticated === 'function') {
                return window.authManager.isAuthenticated();
            }
            if (window.authManager && window.authManager.currentUser) return true;
        } catch (e) { /* ignore */ }

        const token = localStorage.getItem('authToken') || localStorage.getItem('token');
        if (token) return true;

        try {
            const fd = localStorage.getItem('finDeckAuth');
            if (fd) {
                const parsed = JSON.parse(fd);
                if (parsed && (parsed.token || parsed.user)) return true;
            }
        } catch (e) { /* ignore */ }

        return false;
    }

    // Central handler for plan CTA clicks — enforces login and resumes after login
    handleChoosePlan(planKey, approvalUrl) {
        if (!approvalUrl) return;
        if (this.isAuthenticated()) {
            // Open approval link in new tab for authenticated users
            window.open(approvalUrl, '_blank');
            return;
        }

        // Save pending subscription intent and redirect to login
        try {
            localStorage.setItem('pendingSubscription', JSON.stringify({ plan: planKey, approvalUrl, ts: Date.now() }));
        } catch (e) { /* ignore localStorage errors */ }

        // Redirect to login page and return current location so we can resume
        const next = window.location.href;
        window.location.href = `login.html?next=${encodeURIComponent(next)}`;
    }


    // -----------------
    // Billing / Plans
    // -----------------

    async loadSubscriptionInfo() {
        // Guard
        if (!this.apiService || typeof this.apiService.getUserSubscription !== 'function') {
            console.warn('Billing: API service not available or method missing');
            return;
        }

        try {
            const sub = await this.apiService.getUserSubscription();
            console.log('Billing: fetched subscription', sub);
            this.subscriptionData = sub && (sub.data || sub) || null;
            this.renderSubscription(sub);

            // Optionally fetch invoices / billing history
            if (typeof this.apiService.getUserUsage === 'function') {
                const usage = await this.apiService.getUserUsage();
                console.log('Billing: fetched usage', usage);
                // usage may contain invoices or usage metrics depending on API
                this.renderInvoices(usage.invoices || usage.billing || []);
            }
        } catch (err) {
            console.warn('Billing: failed to load subscription info', err);
        }
    }

    renderSubscription(sub) {

        if (!sub) return;

        // Sub might be nested under data
        const data = sub.data || sub;
        const planNameEl = document.getElementById('currentPlanName');
        const planPriceEl = document.getElementById('currentPlanPrice');
        const planPeriodEl = document.getElementById('currentPlanPeriod');
        const planTimeEl = document.getElementById('planTimeRemainingDays');
        const planFeaturesEl = document.getElementById('planFeatures');
        const planBadgeEl = document.getElementById('planBadge');

        // Map backend plan keys to UI labels and prices
        const planMap = {
            free: { label: 'Free', price: 0, period: '/month', credits: 1, features: ['1 conversion'] },
            basic: { label: 'Basic', price: 25, period: '/month', credits: 5, features: ['5 conversions', 'Standard templates'] },
            pro: { label: 'Pro', price: 49, period: '/month', credits: 15, features: ['15 conversions', 'Premium templates', 'Brand kit'] },
            ai: { label: 'AI Pro', price: 99, period: '/month', credits: 'unlimited', features: ['Unlimited conversions', 'AI features', 'Priority support'] }
        };

    // Determine key: prefer explicit plan_key, plan_name, name, or plan/subscription_plan and normalize it
    let rawKey = (data.plan_key || data.plan_name || data.name || data.plan || data.subscription_plan || '') || '';
        let planKey = ('' + rawKey).toLowerCase().trim();

        // Normalize common separators and remove non-alphanumeric
        planKey = planKey.replace(/[_\-\s]+/g, ''); // remove underscores/hyphens/spaces
        planKey = planKey.replace(/[^a-z0-9]/g, ''); // remove any remaining non-alphanum

        // Handle common aliases
        const aiAliases = new Set(['ai', 'aipro', 'ai_pro', 'ai-pro', 'ai pro', 'ai_pro', 'aip', 'aipror', 'aiproplan']);
        if (!planKey || !planMap[planKey]) {
            if (aiAliases.has(planKey) || /^(ai|aipro|aip)/.test(rawKey.toLowerCase())) {
                planKey = 'ai';
            } else if (/^(pro|proplan)$/.test(planKey)) {
                planKey = 'pro';
            } else if (/^(basic|starter|standard)$/.test(planKey)) {
                planKey = 'basic';
            } else if (/^(free|trial)$/.test(planKey)) {
                planKey = 'free';
            } else {
                // default to ai to show AI Pro if ambiguous
                planKey = 'ai';
            }
        }

        const mapped = planMap[planKey] || planMap.ai;

        if (planNameEl) planNameEl.textContent = mapped.label;

        // Update mainpage-specific header and dropdown plan labels
        const headerPlanText = mapped.label || '';
        try {
            const mainHeaderPlan = document.getElementById('userPlan');
            const mainDropdownPlan = document.getElementById('dropdownUserPlan');
            const out = headerPlanText + (headerPlanText.toLowerCase().includes('plan') ? '' : ' Plan');
            if (mainHeaderPlan) mainHeaderPlan.textContent = out;
            if (mainDropdownPlan) mainDropdownPlan.textContent = out;
        } catch (err) { /* ignore DOM issues */ }
        if (planPriceEl) planPriceEl.textContent = `$${mapped.price}`;
        if (planPeriodEl) planPeriodEl.textContent = mapped.period;
        if (planTimeEl) planTimeEl.textContent = data.time_remaining || data.days_until_renewal || 'N/A';
        if (planBadgeEl) planBadgeEl.textContent = data.status ? (data.status === 'active' ? 'Current' : data.status) : 'Current';

        // Render mapped features
        if (planFeaturesEl) {
            planFeaturesEl.innerHTML = '';
            (mapped.features || []).forEach(f => {
                const div = document.createElement('div');
                div.className = 'feature';
                div.textContent = `✓ ${f}`;
                planFeaturesEl.appendChild(div);
            });
        }

    // Payment method
        if (data.payment_method) {
            const pmTitle = document.getElementById('paymentMethodTitle');
            const pmExpiry = document.getElementById('paymentMethodExpiry');
            if (pmTitle) pmTitle.textContent = data.payment_method.brand ? `${this.capitalize(data.payment_method.brand)} ending in ${data.payment_method.last4}` : data.payment_method.description || 'Card on file';
            if (pmExpiry) pmExpiry.textContent = data.payment_method.expiry ? `Expires ${data.payment_method.expiry}` : '';
        }
        // Update price display more explicitly if pricing details exist
        if (data.price || data.amount || data.currency) {
            const priceEl = document.getElementById('currentPlanPrice');
            const periodEl = document.getElementById('currentPlanPeriod');
            if (priceEl) priceEl.textContent = data.price ? `$${data.price}` : (data.amount ? `${data.currency || '$'}${data.amount}` : priceEl.textContent);
            if (periodEl && data.billing_interval) periodEl.textContent = data.billing_interval;
        }
    }

    initSubscriptionButtons() {
        const updateBtn = document.getElementById('updateSubscriptionBtn');
        const updatePaymentBtn = document.getElementById('updatePaymentBtn');

        if (updateBtn) {
            updateBtn.addEventListener('click', (e) => {
                e.preventDefault();
                // Redirect user to the pricing section where PayPal links live
                try {
                    // If already on mainpage, scroll to section smoothly
                    const isOnMain = window.location.pathname.includes('mainpage');
                    if (isOnMain) {
                        const el = document.getElementById('pricing') || document.querySelector('#pricing');
                        if (el && typeof el.scrollIntoView === 'function') {
                            el.scrollIntoView({ behavior: 'smooth', block: 'start' });
                            return;
                        }
                    }
                } catch (err) { /* ignore */ }

                // Otherwise navigate to the landing page pricing anchor
                window.location.href = 'index.html#pricing';
            });
        }

        if (updatePaymentBtn) {
            updatePaymentBtn.addEventListener('click', (e) => {
                e.preventDefault();
                // Same behavior as updateSubscriptionBtn: navigate to pricing where payment links exist
                try {
                    const isOnMain = window.location.pathname.includes('mainpage');
                    if (isOnMain) {
                        const el = document.getElementById('pricing') || document.querySelector('#pricing');
                        if (el && typeof el.scrollIntoView === 'function') {
                            el.scrollIntoView({ behavior: 'smooth', block: 'start' });
                            return;
                        }
                    }
                } catch (err) { /* ignore */ }
                window.location.href = 'index.html#pricing';
            });
        }


        // Remove any previously appended plan subscribe containers (cleanup) so no QR widgets appear site-wide
        try {
            document.querySelectorAll('.plan-subscribe-container, .subscribe-widget, [id^="plan-subscribe-"], [id^="paypal-qr-"]').forEach(el => {
                if (el && el.parentNode) el.parentNode.removeChild(el);
            });
        } catch (e) { /* ignore */ }

        // Only wire CTAs to open approval URLs (enforcing login via handleChoosePlan). Do NOT inject QR widgets.
        try {
            const mapping = {
                basic: 'https://www.paypal.com/ncp/payment/96JM3XH5CLV62',
                pro: 'https://www.paypal.com/ncp/payment/FRZX8QMPTJ93E',
                ai: 'https://www.paypal.com/ncp/payment/M9JQCX6DJHHSJ'
            };

            const selGroups = {
                basic: ['.choose-basic', '#chooseBasicBtn', '[data-plan="basic"]'],
                pro: ['.choose-pro', '#chooseProBtn', '[data-plan="pro"]'],
                ai: ['.choose-ai', '#chooseAiBtn', '[data-plan="ai"]']
            };

            Object.keys(selGroups).forEach(k => {
                selGroups[k].forEach(sel => {
                    document.querySelectorAll(sel).forEach(el => {
                        // Remove any old click handlers by cloning
                        const newEl = el.cloneNode(true);
                        el.parentNode.replaceChild(newEl, el);
                        newEl.addEventListener('click', (ev) => {
                            ev.preventDefault();
                            this.handleChoosePlan(k, mapping[k]);
                        });
                    });
                });
            });
        } catch (e) {
            console.warn('Failed to wire plan CTAs', e);
        }
    }

    openBillingPortal() {
        const sub = this.subscriptionData || null;
        // Check common fields for portal/manage URL
        const portalUrl = sub && (sub.portal_url || sub.manage_url || (sub.data && (sub.data.portal_url || sub.data.manage_url)));
        if (portalUrl) {
            window.open(portalUrl, '_blank');
            return;
        }

        // If no portal URL, fall back to open billing modal
        if (document.getElementById('billingModal')) {
            this.openModal('billingModal');
            return;
        }

        // Final fallback: show notification
        this.showNotification('Billing', 'Manage your subscription from your account dashboard or contact support.', 'info');
    }

    renderInvoices(invoices) {
        const invoiceList = document.getElementById('invoiceList');
        if (!invoiceList) return;

        if (!invoices || invoices.length === 0) {
            invoiceList.innerHTML = `<div class="invoice-empty">No billing history available.</div>`;
            return;
        }

        invoiceList.innerHTML = '';
        invoices.forEach(inv => {
            const item = document.createElement('div');
            item.className = 'invoice-item';
            const date = inv.date || inv.paid_at || inv.created_at || 'Unknown';
            const desc = inv.description || `${inv.plan_name || ''} - $${inv.amount || inv.total || '0.00'}`;
            const status = inv.status || (inv.paid ? 'Paid' : 'Pending');

            item.innerHTML = `
                <div class="invoice-info">
                    <h4>${date}</h4>
                    <p>${desc}</p>
                </div>
                <div class="invoice-actions">
                    <span class="status ${status.toLowerCase()}">${status}</span>
                    ${inv.download_url ? `<button class="neu-button-small" onclick="window.open('${inv.download_url}', '_blank')">Download</button>` : ''}
                </div>
            `;

            invoiceList.appendChild(item);
        });
    }

    capitalize(s) {
        if (!s) return s;
        return s.charAt(0).toUpperCase() + s.slice(1);
    }

    updateCustomAvatar(avatarUrl) {
        // Replace default SVG with custom image if user uploads one
        const profileAvatarContainer = document.querySelector('.profile-avatar-container');
        const userAvatarContainer = document.querySelector('.user-avatar-container');

        if (profileAvatarContainer) {
            profileAvatarContainer.innerHTML = `<img src="${avatarUrl}" alt="Profile" class="profile-avatar">`;
        }
        
        if (userAvatarContainer) {
            userAvatarContainer.innerHTML = `<img src="${avatarUrl}" alt="Profile" class="user-avatar">`;
        }
    }

    /**
     * Fetch quick stats (conversions done / left) and populate profile dropdown
     */
    async fetchAndPopulateStats() {
        // Elements in the profile dropdown
        const doneEl = document.getElementById('statConversionsDone');
        const leftEl = document.getElementById('statConversionsLeft');

        if (!doneEl || !leftEl) return;

        // If user is not authenticated, show a hint and skip the call
        const token = localStorage.getItem('authToken');
        if (!token) {
            doneEl.textContent = 'Login';
            leftEl.textContent = 'Login';
            console.log('fetchAndPopulateStats: no auth token present - skipping stats fetch');
            return;
        }

        // Show loading placeholders
        doneEl.textContent = '…';
        leftEl.textContent = '…';

        if (!this.apiService || typeof this.apiService.getUserUsage !== 'function') {
            // fallback: try to call apiConfig directly
            try {
                const url = (window.apiConfig && window.apiConfig.endpoints && window.apiConfig.endpoints.users && window.apiConfig.endpoints.users.usage) || '/api/v1/users/stats';
                console.log('fetchAndPopulateStats: calling fallback URL', url);
                const resp = await window.apiConfig.makeRequest(url);
                console.log('fetchAndPopulateStats: raw response', resp);
                const data = resp && (resp.data || resp) || {};
                // Prefer presentations_* fields (backend uses these names), then fall back to older names
                const done = data.presentations_created ?? data.presentationsCreated ?? data.conversions_done ?? data.conversionsDone ?? data.done ?? data.used ?? data.slides_created ?? 0;
                const leftFromPresentations = (data.presentations_limit !== undefined && data.presentations_limit !== null) ? (data.presentations_limit - (done || 0)) : undefined;
                const left = data.presentations_left ?? data.presentationsLeft ?? leftFromPresentations ?? data.conversions_left ?? data.conversionsLeft ?? data.remaining ?? data.left ?? 'Unlimited';
                console.log('fetchAndPopulateStats: mapped values', { done, left });
                doneEl.textContent = (done === null || done === undefined) ? '—' : done;
                leftEl.textContent = (left === null || left === undefined) ? '—' : left;
            } catch (err) {
                console.warn('Quick stats (direct) failed', err);
                doneEl.textContent = '—';
                leftEl.textContent = '—';
            }
            return;
        }

        try {
            const usage = await this.apiService.getUserUsage();
            const data = usage && (usage.data || usage) || {};

            // Support multiple possible field names from backend; prefer presentations_* fields first
            const done = data.presentations_created ?? data.presentationsCreated ?? data.conversions_done ?? data.conversionsDone ?? data.done ?? data.used ?? data.slides_created ?? 0;
            const leftFromPresentations = (data.presentations_limit !== undefined && data.presentations_limit !== null) ? (data.presentations_limit - (done || 0)) : undefined;
            const left = data.presentations_left ?? data.presentationsLeft ?? leftFromPresentations ?? data.conversions_left ?? data.conversionsLeft ?? data.remaining ?? data.left ?? (data.limit !== undefined && data.limit !== null ? (data.limit - (done || 0)) : 'Unlimited');

            doneEl.textContent = (done === null || done === undefined) ? '—' : done;
            leftEl.textContent = (left === null || left === undefined) ? '—' : left;
        } catch (err) {
            console.warn('Failed to fetch user usage stats', err);
            doneEl.textContent = '—';
            leftEl.textContent = '—';
        }
    }

    toggleProfileDropdown() {
        const dropdown = document.getElementById('profileDropdown');
        const profileBtn = document.getElementById('profileBtn');
        
        if (!dropdown || !profileBtn) return;

        const isVisible = dropdown.style.display === 'block';
        
        if (isVisible) {
            this.closeProfileDropdown();
        } else {
            this.openProfileDropdown();
        }
    }

    openProfileDropdown() {
        const dropdown = document.getElementById('profileDropdown');
        const profileBtn = document.getElementById('profileBtn');
        
        if (!dropdown || !profileBtn) return;

        dropdown.style.display = 'block';
        setTimeout(() => {
            dropdown.classList.add('show');
        }, 10);
        profileBtn.setAttribute('aria-expanded', 'true');
        // Fetch and populate quick statistics when the dropdown opens
        try {
            this.fetchAndPopulateStats();
        } catch (err) {
            console.warn('Failed to load quick stats', err);
        }
    }

    closeProfileDropdown() {
        const dropdown = document.getElementById('profileDropdown');
        const profileBtn = document.getElementById('profileBtn');
        
        if (!dropdown || !profileBtn) return;

        dropdown.classList.remove('show');
        profileBtn.setAttribute('aria-expanded', 'false');
        setTimeout(() => {
            dropdown.style.display = 'none';
        }, 300);
    }

    initializeModals() {
        // Initialize all modals
        const modals = document.querySelectorAll('.modal-overlay');
        modals.forEach(modal => {
            modal.addEventListener('click', (e) => {
                if (e.target === modal) {
                    this.closeModal(modal.id);
                }
            });
        });

        // Watch profile modal visibility and populate when shown
        const profileModal = document.getElementById('profileModal');
        if (profileModal && typeof loadAndPopulateProfile === 'function') {
            const obs = new MutationObserver((mutations) => {
                mutations.forEach(m => {
                    if (m.attributeName === 'style' || m.attributeName === 'class') {
                        const disp = window.getComputedStyle(profileModal).display;
                        if (disp !== 'none') {
                            try { loadAndPopulateProfile(); } catch (e) { console.warn('profile load failed', e); }
                        }
                    }
                });
            });
            obs.observe(profileModal, { attributes: true, attributeFilter: ['style', 'class'] });
        }

        // Initialize close buttons
        const closeButtons = document.querySelectorAll('.close-btn, .modal-close');
        closeButtons.forEach(btn => {
            btn.addEventListener('click', (e) => {
                const modal = btn.closest('.modal-overlay');
                if (modal) {
                    this.closeModal(modal.id);
                }
            });
        });
    }

    openModal(modalId) {
        const modal = document.getElementById(modalId);
        if (modal) {
            modal.style.display = 'flex';
            modal.classList.add('fade-in');
            document.body.style.overflow = 'hidden';
            
            // Close profile dropdown
            this.closeProfileDropdown();
        }
    }

    closeModal(modalId) {
        const modal = document.getElementById(modalId);
        if (modal) {
            modal.style.display = 'none';
            modal.classList.remove('fade-in');
            document.body.style.overflow = 'auto';
        }
    }

    closeAllModals() {
        const modals = document.querySelectorAll('.modal-overlay');
        modals.forEach(modal => {
            modal.style.display = 'none';
            modal.classList.remove('fade-in');
        });
        document.body.style.overflow = 'auto';
    }

    initializeProgressTracking() {
        // Simulate progress tracking for demo
        const progressBars = document.querySelectorAll('.progress-bar');
        progressBars.forEach(bar => {
            const targetWidth = bar.dataset.progress || '0%';
            setTimeout(() => {
                bar.style.width = targetWidth;
            }, 500);
        });
    }

    showNotification(title, message, type = 'info') {
        // Create notification element
        const notification = document.createElement('div');
        notification.className = `notification notification-${type}`;
        notification.innerHTML = `
            <div class="notification-content">
                <h4>${title}</h4>
                <p>${message}</p>
            </div>
            <button class="notification-close">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <line x1="18" y1="6" x2="6" y2="18"/>
                    <line x1="6" y1="6" x2="18" y2="18"/>
                </svg>
            </button>
        `;

        // Add to page
        document.body.appendChild(notification);

        // Auto remove after 5 seconds
        setTimeout(() => {
            if (notification.parentNode) {
                notification.remove();
            }
        }, 5000);
    }

    // Conversion functionality
    async startConversion() {
        console.log('🔄 Starting conversion process...');
        console.log('📁 Uploaded files:', this.uploadedFiles);
        
        if (this.uploadedFiles.length === 0) {
            console.log('❌ No files uploaded');
            this.showNotification('Error', 'Please upload at least one file to convert.', 'error');
            return;
        }

        if (!this.apiService) {
            console.log('❌ API service not available');
            this.showNotification('Error', 'API service not available. Please refresh the page.', 'error');
            return;
        }

        console.log('✅ Starting conversion with API service');
        this.updateStepIndicator(4);
        this.showConversionProgress();
        
        // Start real conversion process
        await this.performConversion();
    }

    async performConversion() {
        try {
            console.log('🔄 Starting performConversion...');
            const conversionResults = [];
            
            for (let i = 0; i < this.uploadedFiles.length; i++) {
                const file = this.uploadedFiles[i];
                console.log(`🔄 Converting file ${i + 1}/${this.uploadedFiles.length}:`, file);
                
                this.updateConversionProgress(
                    (i / this.uploadedFiles.length) * 100,
                    `Converting ${file.name}...`
                );
                
                // Convert each file - use the file ID from the upload response
                const fileId = file._id || file.id || file.name;
                console.log('📤 Calling API with file ID:', fileId);
                
                const result = await this.apiService.convertExcelToPPT(fileId);
                console.log('📥 API Response:', result);
                
                // For file downloads, the response is the file itself
                if (result && result.ok !== false) {
                    conversionResults.push({
                        originalFile: file,
                        convertedFile: result
                    });
                } else {
                    throw new Error(`Failed to convert ${file.name}`);
                }
            }
            
            console.log('✅ All conversions completed:', conversionResults);
            this.updateConversionProgress(100, 'Conversion complete!');
            this.conversionResults = conversionResults;
            
            setTimeout(() => {
                this.showConversionResults();
            }, 1000);
            
        } catch (error) {
            console.error('Conversion error:', error);
            this.showConversionError(error.message || 'Failed to convert files. Please try again.');
        }
    }

    showConversionProgress() {
        const resultsSection = document.getElementById('resultsSection');
        if (resultsSection) {
            resultsSection.innerHTML = `
                <div class="conversion-progress">
                    <h3>Converting Files...</h3>
                    <div class="progress-container">
                        <div class="progress-bar" id="conversionProgress"></div>
                    </div>
                    <p class="progress-text" id="conversionProgressText">Processing ${this.uploadedFiles.length} file(s)...</p>
                </div>
            `;
        }
    }

    updateConversionProgress(percentage, message) {
        const progressBar = document.getElementById('conversionProgress');
        const progressText = document.getElementById('conversionProgressText');
        
        if (progressBar) {
            progressBar.style.width = `${percentage}%`;
        }
        
        if (progressText) {
            progressText.textContent = message || `Processing... ${Math.round(percentage)}%`;
        }
    }

    showConversionError(errorMessage) {
        const resultsSection = document.getElementById('resultsSection');
        if (resultsSection) {
            resultsSection.innerHTML = `
                <div class="conversion-error">
                    <div class="error-header">
                        <svg class="error-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                            <circle cx="12" cy="12" r="10"/>
                            <line x1="15" y1="9" x2="9" y2="15"/>
                            <line x1="9" y1="9" x2="15" y2="15"/>
                        </svg>
                        <h3>Conversion Failed</h3>
                        <p>${errorMessage}</p>
                    </div>
                    
                    <div class="error-actions">
                        <button class="neu-button primary" onclick="finDeckApp.startConversion()">
                            Try Again
                        </button>
                        <button class="neu-button secondary" onclick="finDeckApp.resetProcess()">
                            Start Over
                        </button>
                    </div>
                </div>
            `;
        }
        
        this.showNotification('Error', errorMessage, 'error');
    }

    showConversionResults() {
        console.log('🎯 Showing conversion results:', this.conversionResults);
        const resultsSection = document.getElementById('downloadSection');
        if (!resultsSection) {
            console.log('❌ Download section not found');
            return;
        }
        if (!this.conversionResults) {
            console.log('❌ No conversion results available');
            return;
        }

        console.log('✅ Creating download links for', this.conversionResults.length, 'files');
        
        // Show the download section and hide processing
        resultsSection.style.display = 'block';
        const processingSection = document.getElementById('processingSection');
        if (processingSection) {
            processingSection.style.display = 'none';
        }
        
        // Update step indicator to final step
        this.updateStepIndicator(5);
        
        const downloadLinks = this.conversionResults.map((result, index) => {
            const convertedFile = result.convertedFile;
            // Get the original filename safely
            const originalName = result.originalFile?.name || 'converted_file';
            const pptFilename = originalName.replace(/\.[^/.]+$/, '.pptx');
            
            return `
                <div class="download-item-enhanced">
                    <div class="file-preview">
                        <div class="file-icon-wrapper">
                            <svg class="file-icon-ppt" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
                                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                                <polyline points="14,2 14,8 20,8"/>
                                <text x="12" y="16" text-anchor="middle" fill="currentColor" font-size="6" font-weight="bold">PPT</text>
                            </svg>
                        </div>
                        <div class="file-details">
                            <div class="file-name-section">
                                <label class="filename-label">Filename:</label>
                                <div class="filename-input-group">
                                    <input type="text" id="filename-${index}" class="filename-input-enhanced" value="${pptFilename.replace('.pptx', '')}" placeholder="Enter filename" onclick="this.select()">
                                    <span class="file-extension-badge">.pptx</span>
                                </div>
                            </div>
                            <div class="file-meta">
                                <span class="file-size">PowerPoint Presentation</span>
                                <span class="conversion-status">
                                    <svg class="status-icon" viewBox="0 0 16 16" fill="currentColor">
                                        <path d="M8 0a8 8 0 1 1 0 16A8 8 0 0 1 8 0zM7 3a1 1 0 0 0-2 0v3.5L3.5 5a1 1 0 0 0-1.414 1.414L4.5 8.5 2.086 10.914A1 1 0 1 0 3.5 12.328L5 10.828V14a1 1 0 1 0 2 0v-3.172l1.5 1.5a1 1 0 0 0 1.414-1.414L7.5 8.5l2.414-2.414A1 1 0 1 0 8.5 4.672L7 6.172V3z"/>
                                    </svg>
                                    Ready
                                </span>
                            </div>
                        </div>
                    </div>
                    <div class="download-actions">
                        <button class="neu-button-small primary download-btn" onclick="finDeckApp.downloadSingleFile(${index})">
                            <svg class="download-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                                <polyline points="7,10 12,15 17,10"/>
                                <line x1="12" y1="15" x2="12" y2="3"/>
                            </svg>
                            <span>Download</span>
                        </button>
                    </div>
                </div>
            `;
        }).join('');

        resultsSection.innerHTML = `
            <div class="conversion-results-enhanced">
                <div class="results-header-enhanced">
                    <div class="success-animation">
                        <svg class="success-icon-large" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                            <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>
                            <polyline points="22,4 12,14.01 9,11.01"/>
                        </svg>
                    </div>
                    <h3 class="results-title">🎉 Conversion Complete!</h3>
                    <p class="results-subtitle">Successfully converted <strong>${this.uploadedFiles.length} file(s)</strong> to PowerPoint presentations.</p>
                </div>
                
                <div class="download-section-enhanced">
                    <div class="section-header">
                        <h4>📥 Download Your Files</h4>
                        <span class="file-count">${this.conversionResults.length} file(s) ready</span>
                    </div>
                    <div class="download-list-enhanced">
                        ${downloadLinks}
                    </div>
                </div>
                </div>
                
                <div class="results-actions">
                    <button class="neu-button primary download-btn" onclick="finDeckApp.downloadAllResults()">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                            <polyline points="7,10 12,15 17,10"/>
                            <line x1="12" y1="15" x2="12" y2="3"/>
                        </svg>
                        Download All Files
                    </button>
                    
                    <button class="neu-button secondary" onclick="finDeckApp.resetProcess()">
                        Convert More Files
                    </button>
                </div>
            </div>
        `;
        
        this.showNotification('Success', 'Files converted successfully!', 'success');
    }

    async downloadSingleFile(index) {
        if (!this.conversionResults || !this.conversionResults[index]) {
            this.showNotification('Error', 'File not found.', 'error');
            return;
        }

        const result = this.conversionResults[index];
        const convertedFile = result.convertedFile;
        
        try {
            this.showNotification('Download', `Downloading PowerPoint file...`, 'info');
            
            // The convertedFile is a Response object from the conversion API
            if (convertedFile && typeof convertedFile.blob === 'function') {
                const blob = await convertedFile.blob();
                
                // Get custom filename from input field
                const filenameInput = document.getElementById(`filename-${index}`);
                let customFilename = filenameInput ? filenameInput.value.trim() : '';
                
                // Sanitize filename (remove invalid characters)
                customFilename = customFilename.replace(/[<>:"/\\|?*]/g, '');
                
                // Fallback to original filename if input is empty or invalid
                if (!customFilename) {
                    const originalName = result.originalFile?.name || 'converted_file';
                    customFilename = originalName.replace(/\.[^/.]+$/, '');
                }
                
                // Ensure .pptx extension
                const filename = customFilename.endsWith('.pptx') ? customFilename : `${customFilename}.pptx`;
                
                // Create download link
                const url = window.URL.createObjectURL(blob);
                const link = document.createElement('a');
                link.href = url;
                link.download = filename;
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                window.URL.revokeObjectURL(url);
                
                this.showNotification('Success', 'PowerPoint file downloaded successfully!', 'success');
            } else {
                throw new Error('Invalid file response');
            }
            
        } catch (error) {
            console.error('Download error:', error);
            this.showNotification('Error', 'Failed to download file. Please try again.', 'error');
        }
    }

    async downloadAllResults() {
        if (!this.conversionResults || this.conversionResults.length === 0) {
            this.showNotification('Error', 'No files to download.', 'error');
            return;
        }

        this.showNotification('Download', 'Starting downloads...', 'info');
        
        // Download each file
        for (let i = 0; i < this.conversionResults.length; i++) {
            try {
                await this.downloadSingleFile(i);
                // Small delay between downloads
                await new Promise(resolve => setTimeout(resolve, 500));
            } catch (error) {
                console.error(`Failed to download file ${i}:`, error);
            }
        }
    }

    // Legacy method for backward compatibility
    downloadResults() {
        this.downloadAllResults();
    }

    resetProcess() {
        this.uploadedFiles = [];
        this.conversionResults = [];
        this.currentStep = 1;
        this.updateUIState();
        this.showNotification('Info', 'Process reset. You can upload new files.', 'info');
    }
}

// --- Profile modal population helpers ---
function populateProfileModalFromUser(user) {
    if (!user) return;
    try {
        const [first, ...rest] = (user.name || '').split(' ');
        const last = rest.join(' ') || '';
        document.getElementById('profileFirstName').value = first || '';
        document.getElementById('profileLastName').value = last || '';
        document.getElementById('profileEmail').value = user.email || '';
        // jobTitle / company / timezone removed from UI (kept in user object if needed server-side)
        // Update avatar image if element exists
        const avatarEl = document.querySelector('.profile-image');
        if (avatarEl) {
            avatarEl.src = user.avatarUrl || 'assets/img/default-avatar.svg';
        }
    } catch (err) {
        console.warn('Failed to populate profile modal:', err);
    }
}

function loadAndPopulateProfile() {
    // Prefer authManager if available
    if (window.authManager && typeof window.authManager.getCurrentUser === 'function') {
        const user = window.authManager.getCurrentUser();
        if (user) {
            populateProfileModalFromUser(user);
            return;
        }
    }

    // Fallback to localStorage
    try {
        const userStr = localStorage.getItem('user');
        if (userStr) {
            const user = JSON.parse(userStr);
            populateProfileModalFromUser(user);
            return;
        }
    } catch (err) {
        console.warn('No user in localStorage or failed to parse');
    }

    // Optionally fetch from API if APIService exists
    if (window.APIService) {
        const svc = new window.APIService();
        svc.get('/api/user').then(resp => {
            if (resp && resp.data) populateProfileModalFromUser(resp.data);
        }).catch(() => {});
    }
}

// Open profile modal and ensure it is populated
function openProfileModal() {
    loadAndPopulateProfile();
    const modal = document.getElementById('profileModal');
    if (modal) modal.style.display = 'block';
}

// (Removed redundant short closeModal helper to avoid duplicate definitions)

// Global functions for backward compatibility
function openModal(modalId) {
    // If opening profile modal, ensure it's populated first
    if (modalId === 'profileModal' && typeof loadAndPopulateProfile === 'function') {
        loadAndPopulateProfile();
    }
    if (window.finDeckApp) {
        window.finDeckApp.openModal(modalId);
    }
}

function closeModal(modalId) {
    if (window.finDeckApp) {
        window.finDeckApp.closeModal(modalId);
    }
}

// Save profile changes handler
async function saveProfileChanges(event) {
    try {
        // Collect values from modal inputs
        const first = (document.getElementById('profileFirstName') || {}).value || '';
        const last = (document.getElementById('profileLastName') || {}).value || '';
        const name = [first, last].filter(Boolean).join(' ');
        const email = (document.getElementById('profileEmail') || {}).value || '';
        const payload = {
            // Keep profile minimal: only update name & email from UI
            name: name || undefined,
            email: email || undefined
        };

        // Basic validation
        if (!payload.name || !payload.email) {
            // Use existing notification helper if present
            if (window.finDeckApp && typeof window.finDeckApp.showNotification === 'function') {
                window.finDeckApp.showNotification('Error', 'Name and email are required.', 'error');
            } else {
                alert('Name and email are required.');
            }
            return;
        }

        // Ensure API service exists
        const svc = (window.APIService) ? new window.APIService() : (window.apiService || null);
        if (!svc || typeof svc.updateUserProfile !== 'function') {
            if (window.finDeckApp && typeof window.finDeckApp.showNotification === 'function') {
                window.finDeckApp.showNotification('Error', 'API service not available. Please refresh the page.', 'error');
            } else {
                alert('API service not available. Please refresh the page.');
            }
            return;
        }

        // Call backend to update profile
        const resp = await svc.updateUserProfile(payload);

        // On success, update authManager and UI
        if (resp) {
            const updated = resp; // response is expected to include user data
            // Update authManager currentUser (only name/email)
            if (window.authManager && window.authManager.currentUser) {
                const cur = window.authManager.currentUser;
                const newUser = Object.assign({}, cur, {
                    name: updated.name || payload.name || cur.name,
                    email: updated.email || payload.email || cur.email
                });
                window.authManager.currentUser = newUser;

                // Persist to local finDeckAuth entry if present
                try {
                    const authStr = localStorage.getItem('finDeckAuth');
                    if (authStr) {
                        const ad = JSON.parse(authStr);
                        ad.user = newUser;
                        localStorage.setItem('finDeckAuth', JSON.stringify(ad));
                    }
                } catch (e) { /* ignore localStorage errors */ }

                // Trigger UI update
                if (typeof window.authManager.updateUI === 'function') window.authManager.updateUI();
                window.dispatchEvent(new CustomEvent('authStateChanged', { detail: { isLoggedIn: true, user: newUser } }));
            }

            // Show success and close modal
            if (window.finDeckApp && typeof window.finDeckApp.showNotification === 'function') {
                window.finDeckApp.showNotification('Success', 'Profile updated successfully', 'success');
            } else {
                alert('Profile updated successfully');
            }
            closeModal('profileModal');
        }
    } catch (err) {
        console.error('Failed to save profile changes', err);
        if (window.finDeckApp && typeof window.finDeckApp.showNotification === 'function') {
            window.finDeckApp.showNotification('Error', err.message || 'Failed to update profile', 'error');
        } else {
            alert('Failed to update profile: ' + (err.message || err));
        }
    }
}

// Wire up save button on DOMContentLoaded so it is available when modal is shown
document.addEventListener('DOMContentLoaded', function() {
    const saveBtn = document.getElementById('profileSaveBtn');
    if (saveBtn) {
        saveBtn.addEventListener('click', saveProfileChanges);
    }
    // Wire Change Password button in Account Settings modal
    const accountChangeBtn = document.getElementById('accountChangePasswordBtn');
    if (accountChangeBtn) {
        accountChangeBtn.addEventListener('click', async function(e){
            e.preventDefault();

            // Prefer current user email from authManager, otherwise use stored modal value
            const email = (window.authManager && window.authManager.currentUser && window.authManager.currentUser.email)
                || (document.getElementById('profileEmail') || {}).value || '';

            if (!email) {
                if (window.finDeckApp && typeof window.finDeckApp.showNotification === 'function') {
                    window.finDeckApp.showNotification('Error', 'No email available for this account.', 'error');
                } else {
                    alert('No email available for this account.');
                }
                return;
            }

            accountChangeBtn.disabled = true;
            const originalText = accountChangeBtn.textContent;
            accountChangeBtn.textContent = 'Sending…';

            try {
                const base = window.apiConfig ? window.apiConfig.API_BASE : '';
                const url = base ? `${base}/codes/send` : '/api/v1/codes/send';
                const res = await fetch(url, {
                    method: 'POST',
                    headers: {'Content-Type':'application/json'},
                    body: JSON.stringify({ email, purpose: 'password_reset' })
                });

                let payload = null;
                try { payload = await res.json(); } catch (e) { payload = null; }

                if (res.ok) {
                    if (window.finDeckApp && typeof window.finDeckApp.showNotification === 'function') {
                        window.finDeckApp.showNotification('Success', 'Reset code sent. Redirecting…', 'success');
                    }
                    setTimeout(() => {
                        window.location.href = `reset_password.html?email=${encodeURIComponent(email)}`;
                    }, 600);
                } else {
                    const msg = payload && (payload.detail || payload.message) || 'Failed to send reset code';
                    if (window.finDeckApp && typeof window.finDeckApp.showNotification === 'function') window.finDeckApp.showNotification('Error', msg, 'error');
                    else alert(msg);
                }
            } catch (err) {
                console.error('Account Change Password request failed', err);
                if (window.finDeckApp && typeof window.finDeckApp.showNotification === 'function') window.finDeckApp.showNotification('Error', 'Network error while sending reset code', 'error');
                else alert('Network error while sending reset code');
            } finally {
                accountChangeBtn.disabled = false;
                accountChangeBtn.textContent = originalText;
            }
        });
    }
    // Wire Delete Account button
    const deleteBtn = document.getElementById('deleteAccountBtn');
    if (deleteBtn) {
        deleteBtn.addEventListener('click', async function(e){
            e.preventDefault();

            // Double-confirm destructive action
            const ok = confirm('Delete account? This will deactivate your account and cannot be undone. Are you sure?');
            if (!ok) return;

            // Ensure we have an auth token
            const token = localStorage.getItem('authToken');
            if (!token) {
                if (window.finDeckApp && typeof window.finDeckApp.showNotification === 'function') {
                    window.finDeckApp.showNotification('Error', 'You must be signed in to delete your account.', 'error');
                } else {
                    alert('You must be signed in to delete your account.');
                }
                return;
            }

            // disable button while working
            deleteBtn.disabled = true;
            const orig = deleteBtn.textContent;
            deleteBtn.textContent = 'Deleting…';

            try {
                const base = window.apiConfig ? window.apiConfig.API_BASE : '';
                const url = base ? `${base}/users/me` : '/api/v1/users/me';
                const res = await fetch(url, { method: 'DELETE', headers: { 'Authorization': `Bearer ${token}` } });
                let body = null;
                try { body = await res.json(); } catch (e) { body = null; }

                if (res.ok) {
                    // Clear local session and redirect to signup/login
                    try { localStorage.removeItem('authToken'); localStorage.removeItem('tokenType'); localStorage.removeItem('finDeckAuth'); localStorage.removeItem('currentUser'); } catch(e){}
                    if (window.finDeckApp && typeof window.finDeckApp.showNotification === 'function') window.finDeckApp.showNotification('Success', 'Account deleted. Redirecting...', 'success');
                    setTimeout(()=> { window.location.href = 'index.html'; }, 700);
                } else {
                    const msg = body && (body.detail || body.message) || res.statusText || 'Failed to delete account';
                    if (window.finDeckApp && typeof window.finDeckApp.showNotification === 'function') window.finDeckApp.showNotification('Error', msg, 'error');
                    else alert(msg);
                }
            } catch (err) {
                console.error('Delete account failed', err);
                if (window.finDeckApp && typeof window.finDeckApp.showNotification === 'function') window.finDeckApp.showNotification('Error', 'Network error while deleting account', 'error');
                else alert('Network error while deleting account');
            } finally {
                deleteBtn.disabled = false;
                deleteBtn.textContent = orig;
            }
        });
    }
});

// Ensure subscription update buttons always redirect to pricing even if the app
// instance hasn't been initialized yet (defensive wiring).
document.addEventListener('DOMContentLoaded', function() {
    try {
        const wireRedirect = (btn) => {
            if (!btn) return;
            // Avoid binding twice
            if (btn.__wiredToPricing) return;
            btn.__wiredToPricing = true;
            btn.addEventListener('click', (e) => {
                e.preventDefault();
                try {
                    const isOnMain = window.location.pathname.includes('mainpage');
                    if (isOnMain) {
                        const el = document.getElementById('pricing') || document.querySelector('#pricing');
                        if (el && typeof el.scrollIntoView === 'function') {
                            el.scrollIntoView({ behavior: 'smooth', block: 'start' });
                            return;
                        }
                    }
                } catch (err) { /* ignore */ }
                window.location.href = 'mainpage.html#pricing';
            });
        };

        wireRedirect(document.getElementById('updateSubscriptionBtn'));
        wireRedirect(document.getElementById('updatePaymentBtn'));
    } catch (err) {
        // Defensive: do not break page if something goes wrong
        console.warn('Failed to wire subscription redirect buttons', err);
    }
});

function showTemplateSettings() {
    window.location.href = 'templates.html';
}

function showHelp() {
    window.location.href = 'help.html';
}

// Global function to test conversion
function testConversion() {
    console.log('🧪 Testing conversion...');
    if (window.finDeckApp) {
        window.finDeckApp.startConversion();
    } else {
        console.log('❌ FinDeck app not available');
    }
}

// Initialize the application with proper error handling
document.addEventListener('DOMContentLoaded', function() {
    console.log('🚀 Initializing FinDeck App...');
    // Prevent initializing the full FinDeck app on marketing/landing pages that don't include app elements.
    // If none of the key app elements exist, skip initialization to avoid errors on index.html.
    if (!document.getElementById('uploadArea') && !document.getElementById('profileBtn') && !document.getElementById('convertBtn')) {
        console.log('FinDeckApp: skipping initialization — app elements not present on this page.');
        return;
    }

    // Check if required dependencies are available
    if (typeof window.APIService === 'undefined') {
        console.warn('⚠️ APIService not loaded yet, retrying...');
        setTimeout(() => {
            if (typeof window.APIService !== 'undefined') {
                console.log('✅ APIService loaded successfully');
                window.finDeckApp = new FinDeckApp();
                // Initialize convert button after app is ready
                setTimeout(() => {
                    if (window.finDeckApp) {
                        window.finDeckApp.initializeConvertButton();
                    }
                }, 100);
                // If user had a pending subscription intent (saved before redirect to login), resume it now
                try {
                    const pending = localStorage.getItem('pendingSubscription');
                    if (pending) {
                        const p = JSON.parse(pending);
                        if (p && p.approvalUrl && window.finDeckApp && window.finDeckApp.isAuthenticated && window.finDeckApp.isAuthenticated()) {
                            // Open approval link and clear pending
                            window.open(p.approvalUrl, '_blank');
                            localStorage.removeItem('pendingSubscription');
                        }
                    }
                } catch (e) { /* ignore */ }
            } else {
                console.error('❌ APIService failed to load');
            }
        }, 100);
    } else {
        console.log('✅ All dependencies loaded');
        window.finDeckApp = new FinDeckApp();
        // Initialize convert button after app is ready
        setTimeout(() => {
            if (window.finDeckApp) {
                window.finDeckApp.initializeConvertButton();
            }
        }, 100);
        // Cleanup any previously appended plan subscribe containers/QRs left by older runs
        try {
            // Fix: use valid "starts with" attribute selector (^=) instead of invalid "^^"
            document.querySelectorAll('.plan-subscribe-container, .subscribe-widget, [id^="plan-subscribe-"], [id^="paypal-qr-"], [id^="paypal-hosted-"]').forEach(el => {
                if (el && el.parentNode) el.parentNode.removeChild(el);
            });
        } catch (e) { /* ignore */ }

        // Resume any pending subscription approval after login
        try {
            const pending = localStorage.getItem('pendingSubscription');
            if (pending) {
                const p = JSON.parse(pending);
                if (p && p.approvalUrl && window.finDeckApp && window.finDeckApp.isAuthenticated && window.finDeckApp.isAuthenticated()) {
                    window.open(p.approvalUrl, '_blank');
                    localStorage.removeItem('pendingSubscription');
                }
            }
        } catch (e) { /* ignore */ }
    }
});