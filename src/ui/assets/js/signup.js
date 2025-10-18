// FinDeck Signup Form JavaScript
class NeumorphismSignupForm {
    constructor() {
        console.log('NeumorphismSignupForm constructor called');
        
        this.form = document.getElementById('signupForm');
        this.firstNameInput = document.getElementById('firstName');
        this.lastNameInput = document.getElementById('lastName');
        this.emailInput = document.getElementById('email');
        this.passwordInput = document.getElementById('password');
        this.confirmPasswordInput = document.getElementById('confirmPassword');
        this.passwordToggle = document.getElementById('passwordToggle');
        this.confirmPasswordToggle = document.getElementById('confirmPasswordToggle');
        this.termsCheckbox = document.getElementById('terms');
        this.newsletterCheckbox = document.getElementById('newsletter');
        this.submitButton = this.form ? this.form.querySelector('.signup-btn') : null;
        this.successMessage = document.getElementById('successMessage');
        this.socialButtons = document.querySelectorAll('.neu-social');
        
        // Detailed logging
        console.log('Form elements check:', {
            form: !!this.form,
            firstName: !!this.firstNameInput,
            lastName: !!this.lastNameInput,
            email: !!this.emailInput,
            password: !!this.passwordInput,
            confirmPassword: !!this.confirmPasswordInput,
            terms: !!this.termsCheckbox,
            submitButton: !!this.submitButton,
            successMessage: !!this.successMessage
        });
        
        if (!this.form) {
            console.error('Critical error: Signup form not found!');
            return;
        }
        
        // Check for missing required elements
        const missingElements = [];
        if (!this.firstNameInput) missingElements.push('firstName');
        if (!this.lastNameInput) missingElements.push('lastName');
        if (!this.emailInput) missingElements.push('email');
        if (!this.passwordInput) missingElements.push('password');
        if (!this.confirmPasswordInput) missingElements.push('confirmPassword');
        if (!this.termsCheckbox) missingElements.push('terms');
        if (!this.submitButton) missingElements.push('submitButton');
        
        if (missingElements.length > 0) {
            console.error('Missing form elements:', missingElements);
            alert('Form setup error: Missing elements - ' + missingElements.join(', '));
            return;
        }
        
        console.log('All form elements found, initializing...');
        this.init();
    }
    
    init() {
        this.bindEvents();
        this.setupPasswordToggles();
        this.setupSocialButtons();
        this.setupNeumorphicEffects();
    }
    
    bindEvents() {
        if (!this.form) {
            console.error('Cannot bind events: form not found');
            return;
        }
        
        this.form.addEventListener('submit', (e) => this.handleSubmit(e));
        
        // Input validation events
        if (this.firstNameInput) this.firstNameInput.addEventListener('blur', () => this.validateFirstName());
        if (this.lastNameInput) this.lastNameInput.addEventListener('blur', () => this.validateLastName());
        if (this.emailInput) this.emailInput.addEventListener('blur', () => this.validateEmail());
        if (this.passwordInput) this.passwordInput.addEventListener('blur', () => this.validatePassword());
        if (this.confirmPasswordInput) this.confirmPasswordInput.addEventListener('blur', () => this.validateConfirmPassword());
        if (this.termsCheckbox) this.termsCheckbox.addEventListener('change', () => this.validateTerms());
        
        // Clear errors on input
        if (this.firstNameInput) this.firstNameInput.addEventListener('input', () => this.clearError('firstName'));
        if (this.lastNameInput) this.lastNameInput.addEventListener('input', () => this.clearError('lastName'));
        if (this.emailInput) this.emailInput.addEventListener('input', () => this.clearError('email'));
        if (this.passwordInput) {
            this.passwordInput.addEventListener('input', () => {
                this.clearError('password');
                if (this.confirmPasswordInput && this.confirmPasswordInput.value) {
                    this.validateConfirmPassword();
                }
            });
        }
        if (this.confirmPasswordInput) this.confirmPasswordInput.addEventListener('input', () => this.clearError('confirmPassword'));
        
        // Add soft press effects to inputs
        const inputs = [this.firstNameInput, this.lastNameInput, this.emailInput, 
                       this.passwordInput, this.confirmPasswordInput].filter(input => input);
        inputs.forEach(input => {
            input.addEventListener('focus', (e) => this.addSoftPress(e));
            input.addEventListener('blur', (e) => this.removeSoftPress(e));
        });
    }
    
    setupPasswordToggles() {
        // Main password toggle
        this.passwordToggle.addEventListener('click', () => {
            const type = this.passwordInput.type === 'password' ? 'text' : 'password';
            this.passwordInput.type = type;
            this.passwordToggle.classList.toggle('show-password', type === 'text');
            this.animateSoftPress(this.passwordToggle);
        });
        
        // Confirm password toggle
        this.confirmPasswordToggle.addEventListener('click', () => {
            const type = this.confirmPasswordInput.type === 'password' ? 'text' : 'password';
            this.confirmPasswordInput.type = type;
            this.confirmPasswordToggle.classList.toggle('show-password', type === 'text');
            this.animateSoftPress(this.confirmPasswordToggle);
        });
    }
    
    setupSocialButtons() {
        this.socialButtons.forEach(button => {
            button.addEventListener('click', (e) => {
                e.preventDefault();
                const provider = button.getAttribute('data-provider');
                this.handleSocialSignup(provider, button);
            });
        });
    }
    
    setupNeumorphicEffects() {
        // Add hover effects to all neumorphic elements
        const neuElements = document.querySelectorAll('.neu-icon, .neu-checkbox, .neu-social');
        neuElements.forEach(element => {
            element.addEventListener('mouseenter', () => {
                element.style.transform = 'scale(1.05)';
            });
            
            element.addEventListener('mouseleave', () => {
                element.style.transform = 'scale(1)';
            });
        });
    }
    
    addSoftPress(e) {
        const inputGroup = e.target.closest('.neu-input');
        inputGroup.style.transform = 'scale(0.98)';
    }
    
    removeSoftPress(e) {
        const inputGroup = e.target.closest('.neu-input');
        inputGroup.style.transform = 'scale(1)';
    }
    
    animateSoftPress(element) {
        element.style.transform = 'scale(0.95)';
        setTimeout(() => {
            element.style.transform = 'scale(1)';
        }, 150);
    }
    
    handleSubmit(e) {
        e.preventDefault();
        console.log('Form submission intercepted successfully');
        
        // Double-check form elements exist
        if (!this.firstNameInput || !this.lastNameInput || !this.emailInput || 
            !this.passwordInput || !this.confirmPasswordInput || !this.termsCheckbox) {
            console.error('Form elements missing during submit');
            alert('Form error: Some form elements are missing');
            return;
        }
        
        // Get current values for debugging
        const formData = {
            firstName: this.firstNameInput.value.trim(),
            lastName: this.lastNameInput.value.trim(),
            email: this.emailInput.value.trim(),
            password: this.passwordInput.value,
            confirmPassword: this.confirmPasswordInput.value,
            termsChecked: this.termsCheckbox.checked
        };
        
        console.log('Form data at submit:', formData);
        
        // Validate all fields
        console.log('Starting validation...');
        const isFirstNameValid = this.validateFirstName();
        const isLastNameValid = this.validateLastName();
        const isEmailValid = this.validateEmail();
        const isPasswordValid = this.validatePassword();
        const isConfirmPasswordValid = this.validateConfirmPassword();
        const isTermsValid = this.validateTerms();
        
        console.log('Validation results:', {
            firstName: isFirstNameValid,
            lastName: isLastNameValid,
            email: isEmailValid,
            password: isPasswordValid,
            confirmPassword: isConfirmPasswordValid,
            terms: isTermsValid
        });
        
        // Show validation results to user if any field fails
        if (!isFirstNameValid || !isLastNameValid || !isEmailValid || 
            !isPasswordValid || !isConfirmPasswordValid || !isTermsValid) {
            
            const issues = [];
            if (!isFirstNameValid) issues.push('First name is required');
            if (!isLastNameValid) issues.push('Last name is required');
            if (!isEmailValid) issues.push('Valid email address is required');
            if (!isPasswordValid) issues.push('Password must be at least 6 characters');
            if (!isConfirmPasswordValid) issues.push('Passwords must match');
            if (!isTermsValid) issues.push('You must accept the terms and conditions');
            
            alert('Please fix the following issues:\n\n' + issues.join('\n'));
            console.log('Validation failed with issues:', issues);
            return;
        }
        
        console.log('All validations passed! Calling submitForm...');
        this.submitForm();
    }
    
    validateFirstName() {
        const firstName = this.firstNameInput.value.trim();
        if (!firstName) {
            this.showError('firstName', 'First name is required');
            return false;
        }
        if (firstName.length < 2) {
            this.showError('firstName', 'First name must be at least 2 characters');
            return false;
        }
        this.clearError('firstName');
        return true;
    }
    
    validateLastName() {
        const lastName = this.lastNameInput.value.trim();
        if (!lastName) {
            this.showError('lastName', 'Last name is required');
            return false;
        }
        if (lastName.length < 2) {
            this.showError('lastName', 'Last name must be at least 2 characters');
            return false;
        }
        this.clearError('lastName');
        return true;
    }
    
    validateEmail() {
        const email = this.emailInput.value.trim();
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        
        if (!email) {
            this.showError('email', 'Email address is required');
            return false;
        }
        if (!emailRegex.test(email)) {
            this.showError('email', 'Please enter a valid email address');
            return false;
        }
        this.clearError('email');
        return true;
    }
    
    validatePassword() {
        const password = this.passwordInput.value;
        const minLength = 6; // Reduced from 8
        
        if (!password) {
            this.showError('password', 'Password is required');
            return false;
        }
        if (password.length < minLength) {
            this.showError('password', `Password must be at least ${minLength} characters long`);
            return false;
        }
        
        // Simplified validation - only require length for now
        // TODO: Add back complexity requirements for production
        
        this.clearError('password');
        return true;
    }
    
    validateConfirmPassword() {
        const password = this.passwordInput.value;
        const confirmPassword = this.confirmPasswordInput.value;
        
        if (!confirmPassword) {
            this.showError('confirmPassword', 'Please confirm your password');
            return false;
        }
        if (password !== confirmPassword) {
            this.showError('confirmPassword', 'Passwords do not match');
            return false;
        }
        this.clearError('confirmPassword');
        return true;
    }
    
    validateTerms() {
        if (!this.termsCheckbox.checked) {
            this.showError('terms', 'You must agree to the Terms of Service and Privacy Policy');
            return false;
        }
        this.clearError('terms');
        return true;
    }
    
    showError(fieldName, message) {
        try {
            const fieldElement = document.getElementById(fieldName);
            const errorElement = document.getElementById(fieldName + 'Error');
            
            if (!fieldElement) {
                console.error(`Field element not found: ${fieldName}`);
                return;
            }
            
            if (!errorElement) {
                console.error(`Error element not found: ${fieldName}Error`);
                return;
            }
            
            const formGroup = fieldElement.closest('.form-group');
            if (formGroup) {
                formGroup.classList.add('error');
            }
            
            errorElement.textContent = message;
            console.log(`Error shown for ${fieldName}: ${message}`);
        } catch (error) {
            console.error(`Error in showError for ${fieldName}:`, error);
        }
    }
    
    clearError(fieldName) {
        try {
            const fieldElement = document.getElementById(fieldName);
            const errorElement = document.getElementById(fieldName + 'Error');
            
            if (!fieldElement || !errorElement) {
                return; // Silently fail if elements don't exist
            }
            
            const formGroup = fieldElement.closest('.form-group');
            if (formGroup) {
                formGroup.classList.remove('error');
            }
            
            errorElement.textContent = '';
        } catch (error) {
            console.error(`Error in clearError for ${fieldName}:`, error);
        }
    }
    
    async submitForm() {
        console.log('submitForm called - starting signup process');
        
        // Double-check all form elements before proceeding
        if (!this.firstNameInput || !this.lastNameInput || !this.emailInput || 
            !this.passwordInput || !this.confirmPasswordInput) {
            console.error('Form elements missing in submitForm');
            alert('Form error: Missing required elements');
            return;
        }
        
        // Show loading state
        if (this.submitButton) {
            this.submitButton.classList.add('loading');
            console.log('Loading state activated');
        }
        
        try {
            console.log('Starting real API signup...');
            
            // Check if API service is available
            if (!window.apiService) {
                throw new Error('API service not loaded. Please refresh the page and try again.');
            }
            
            // Get form data
            const firstName = this.firstNameInput.value.trim();
            const lastName = this.lastNameInput.value.trim();
            const email = this.emailInput.value.trim();
            const password = this.passwordInput.value;
            
            // Validate data one more time
            if (!firstName || !lastName || !email || !password) {
                throw new Error('Invalid form data: missing required fields');
            }
            
            // Prepare user data for API
            const userData = {
                name: `${firstName} ${lastName}`,
                email: email,
                password: password
            };
            
            console.log('Calling real API signup...');
            
            // Call real API signup
            const response = await window.apiService.signup(userData);
            console.log('API signup successful:', response);
            
            // Create user data for local auth manager
            const localUserData = {
                name: userData.name,
                firstName: firstName,
                lastName: lastName,
                email: email,
                plan: 'Free',
                conversions: 0,
                templates: 0,
                loginTime: new Date().toISOString(),
                id: response.id || 'temp-id'
            };
            
            console.log('User data created:', localUserData);
            
            // Check if auth manager exists and is properly initialized
            if (!window.authManager) {
                console.error('Auth manager not found!');
                throw new Error('Authentication system not available');
            }
            
            console.log('Auth manager found, logging in user...');
            const loginResult = window.authManager.login(userData);
            console.log('Login result:', loginResult);
            
            // Verify login was successful
            if (!window.authManager.isLoggedIn) {
                throw new Error('Login failed - user not logged in after signup');
            }
            
            console.log('User logged in successfully');
            
            // Hide form and show success message
            if (this.form) {
                this.form.style.display = 'none';
                console.log('Form hidden');
            }
            if (this.successMessage) {
                this.successMessage.classList.add('show');
                console.log('Success message shown');
            }
            
            console.log('Starting redirect process...');
            
            // Redirect to main page after success
            setTimeout(() => {
                console.log('Executing redirect to mainpage.html...');
                
                // Try multiple redirect methods
                try {
                    window.location.href = 'mainpage.html';
                } catch (redirectError) {
                    console.error('Primary redirect failed:', redirectError);
                    try {
                        window.location.replace('mainpage.html');
                    } catch (fallbackError) {
                        console.error('Fallback redirect failed:', fallbackError);
                        window.location.assign('mainpage.html');
                    }
                }
            }, 1500);
            
        } catch (error) {
            console.error('Signup error details:', error);
            alert('Signup failed: ' + error.message);
            this.showError('email', 'An error occurred during signup: ' + error.message);
        } finally {
            if (this.submitButton) {
                this.submitButton.classList.remove('loading');
                console.log('Loading state removed');
            }
        }
    }
    
    async simulateSignup() {
        // Simulate network delay
        return new Promise((resolve, reject) => {
            setTimeout(() => {
                // Always succeed for demo (remove random failure)
                resolve({ success: true, message: 'Account created successfully' });
            }, 2000);
        });
    }
    
    handleSocialSignup(provider, button) {
        console.log(`Signup with ${provider}`);
        
        // Add click animation
        this.animateSoftPress(button);
        
        // Simulate social signup process
        const originalText = button.innerHTML;
        button.style.opacity = '0.7';
        
        setTimeout(() => {
            button.style.opacity = '1';
            console.log(`${provider} signup completed`);
            // In a real app, integrate with social auth providers
        }, 1000);
    }
}

// Initialize the signup form when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    console.log('DOM loaded, initializing signup form...');
    try {
        const form = new NeumorphismSignupForm();
        console.log('Signup form initialized successfully');
    } catch (error) {
        console.error('Error initializing signup form:', error);
    }
});