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

    handleFiles(files) {
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
            this.uploadedFiles = [...this.uploadedFiles, ...validFiles];
            this.displayUploadedFiles();
            this.updateStepIndicator(2);
            this.showFileListSection();
        }
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
    startConversion() {
        if (this.uploadedFiles.length === 0) {
            this.showNotification('Error', 'Please upload at least one file to convert.', 'error');
            return;
        }

        this.updateStepIndicator(4);
        this.showConversionProgress();
        
        // Simulate conversion process
        this.simulateConversion();
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
                    <p class="progress-text">Processing ${this.uploadedFiles.length} file(s)...</p>
                </div>
            `;
        }
    }

    simulateConversion() {
        const progressBar = document.getElementById('conversionProgress');
        const progressText = document.querySelector('.progress-text');
        
        let progress = 0;
        const interval = setInterval(() => {
            progress += Math.random() * 15;
            
            if (progress >= 100) {
                progress = 100;
                clearInterval(interval);
                this.showConversionResults();
            }
            
            if (progressBar) {
                progressBar.style.width = `${progress}%`;
            }
            
            if (progressText) {
                progressText.textContent = `Processing... ${Math.round(progress)}%`;
            }
        }, 500);
    }

    showConversionResults() {
        const resultsSection = document.getElementById('resultsSection');
        if (resultsSection) {
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
                    
                    <div class="results-actions">
                        <button class="neu-button primary download-btn" onclick="finDeckApp.downloadResults()">
                            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                                <polyline points="7,10 12,15 17,10"/>
                                <line x1="12" y1="15" x2="12" y2="3"/>
                            </svg>
                            Download PowerPoint Files
                        </button>
                        
                        <button class="neu-button secondary" onclick="finDeckApp.resetProcess()">
                            Convert More Files
                        </button>
                    </div>
                </div>
            `;
        }
        
        this.showNotification('Success', 'Files converted successfully!', 'success');
    }

    downloadResults() {
        // Simulate download
        this.showNotification('Download', 'Starting download...', 'info');
        
        // In a real application, you would trigger the actual download
        setTimeout(() => {
            this.showNotification('Success', 'Files downloaded successfully!', 'success');
        }, 2000);
    }

    resetProcess() {
        this.uploadedFiles = [];
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