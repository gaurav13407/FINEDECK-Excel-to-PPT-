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
        
        // Initialize API service
        this.apiService = window.APIService ? new window.APIService() : null;
        
        this.init();
    }

    init() {
        this.initializeEventListeners();
        this.initializeFileUpload();
        this.initializeProfileDropdown();
        this.initializeModals();
        this.initializeStepIndicator();
        this.initializeProgressTracking();
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

        // Click to upload
        uploadArea.addEventListener('click', () => {
            this.openFileDialog();
        });

        // Browse button
        if (browseBtn) {
            browseBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                this.openFileDialog();
            });
        }

        // File input change
        fileInput.addEventListener('change', (e) => {
            const files = Array.from(e.target.files);
            this.handleFiles(files);
        });
    }

    openFileDialog() {
        const fileInput = document.getElementById('fileInput');
        if (fileInput) {
            fileInput.click();
        }
    }

    async handleFiles(files) {
        if (!files || files.length === 0) return;

        const validFiles = [];
        const errors = [];

        files.forEach(file => {
            const validation = this.validateFile(file);
            if (validation.valid) {
                validFiles.push(file);
            } else {
                errors.push({
                    fileName: file.name,
                    error: validation.error
                });
            }
        });

        if (errors.length > 0) {
            this.showFileErrors(errors);
        }

        if (validFiles.length > 0) {
            // Upload files to backend
            await this.uploadFilesToBackend(validFiles);
        }
    }

    async uploadFilesToBackend(files) {
        if (!this.apiService) {
            this.showNotification('Error', 'API service not available. Please refresh the page.', 'error');
            return;
        }

        this.showUploadProgress();

        try {
            for (let i = 0; i < files.length; i++) {
                const file = files[i];
                this.updateUploadProgress(i + 1, files.length, `Uploading ${file.name}...`);
                
                const response = await this.apiService.uploadFile(file);
                
                if (response.success) {
                    // Add file metadata from backend response
                    const fileWithId = {
                        ...file,
                        id: response.data.id,
                        url: response.data.url,
                        uploaded: true
                    };
                    this.uploadedFiles.push(fileWithId);
                } else {
                    throw new Error(response.message || `Failed to upload ${file.name}`);
                }
            }

            this.hideUploadProgress();
            this.displayUploadedFiles();
            this.updateStepIndicator(2);
            this.showFileListSection();
            this.showNotification('Success', `Successfully uploaded ${files.length} file(s)!`, 'success');

        } catch (error) {
            console.error('Upload error:', error);
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

        // Close button functionality
        const closeBtn = notification.querySelector('.notification-close');
        closeBtn.addEventListener('click', () => {
            notification.remove();
        });
    }

    handleResize() {
        // Handle responsive layout changes
        const isMobile = window.innerWidth <= 768;
        
        if (isMobile) {
            this.closeProfileDropdown();
        }
    }

    updateUIState() {
        // Update UI based on current state
        this.updateStepIndicator(this.currentStep);
        
        if (this.uploadedFiles.length > 0) {
            this.displayUploadedFiles();
            this.showFileListSection();
        } else {
            this.showUploadSection();
        }
    }

    // Conversion functionality
    async startConversion() {
        if (this.uploadedFiles.length === 0) {
            this.showNotification('Error', 'Please upload at least one file to convert.', 'error');
            return;
        }

        if (!this.apiService) {
            this.showNotification('Error', 'API service not available. Please refresh the page.', 'error');
            return;
        }

        this.updateStepIndicator(4);
        this.showConversionProgress();
        
        // Start real conversion process
        await this.performConversion();
    }

    async performConversion() {
        try {
            const conversionResults = [];
            
            for (let i = 0; i < this.uploadedFiles.length; i++) {
                const file = this.uploadedFiles[i];
                
                this.updateConversionProgress(
                    (i / this.uploadedFiles.length) * 100,
                    `Converting ${file.name}...`
                );
                
                // Convert each file
                const result = await this.apiService.convertExcelToPPT(file.id || file.name);
                
                if (result.success) {
                    conversionResults.push({
                        originalFile: file,
                        convertedFile: result.data
                    });
                } else {
                    throw new Error(result.message || `Failed to convert ${file.name}`);
                }
            }
            
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
        const resultsSection = document.getElementById('resultsSection');
        if (!resultsSection || !this.conversionResults) return;

        const downloadLinks = this.conversionResults.map((result, index) => {
            const convertedFile = result.convertedFile;
            return `
                <div class="download-item">
                    <div class="file-info">
                        <svg class="file-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                            <polyline points="14,2 14,8 20,8"/>
                        </svg>
                        <span>${convertedFile.filename || result.originalFile.name.replace(/\.[^/.]+$/, '.pptx')}</span>
                    </div>
                    <button class="neu-button-small primary" onclick="finDeckApp.downloadSingleFile(${index})">
                        Download
                    </button>
                </div>
            `;
        }).join('');

        resultsSection.innerHTML = `
            <div class="conversion-results">
                <div class="results-header">
                    <svg class="success-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>
                        <polyline points="22,4 12,14.01 9,11.01"/>
                    </svg>
                    <h3>Conversion Complete!</h3>
                    <p>Successfully converted ${this.uploadedFiles.length} file(s) to PowerPoint.</p>
                </div>
                
                <div class="download-list">
                    ${downloadLinks}
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
            this.showNotification('Download', `Downloading ${convertedFile.filename}...`, 'info');
            
            // Create download link
            if (convertedFile.url) {
                const link = document.createElement('a');
                link.href = convertedFile.url;
                link.download = convertedFile.filename || result.originalFile.name.replace(/\.[^/.]+$/, '.pptx');
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                
                this.showNotification('Success', 'File downloaded successfully!', 'success');
            } else {
                throw new Error('Download URL not available');
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

// Global functions for backward compatibility
function openModal(modalId) {
    if (window.finDeckApp) {
        window.finDeckApp.openModal(modalId);
    }
}

function closeModal(modalId) {
    if (window.finDeckApp) {
        window.finDeckApp.closeModal(modalId);
    }
}

function showTemplateSettings() {
    window.location.href = 'templates.html';
}

function showHelp() {
    window.location.href = 'help.html';
}

// Initialize the application
window.finDeckApp = new FinDeckApp();