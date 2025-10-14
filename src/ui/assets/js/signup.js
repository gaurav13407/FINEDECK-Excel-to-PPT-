// FinDeck Signup Form JavaScript
class NeumorphismSignupForm {
    constructor() {
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
        this.submitButton = this.form.querySelector('.signup-btn');
        this.successMessage = document.getElementById('successMessage');
        this.socialButtons = document.querySelectorAll('.neu-social');
        
        this.init();
    }
    
    init() {
        this.bindEvents();
        this.setupPasswordToggles();
        this.setupSocialButtons();
        this.setupNeumorphicEffects();
    }
    
    bindEvents() {
        this.form.addEventListener('submit', (e) => this.handleSubmit(e));
        
        // Input validation events
        this.firstNameInput.addEventListener('blur', () => this.validateFirstName());
        this.lastNameInput.addEventListener('blur', () => this.validateLastName());
        this.emailInput.addEventListener('blur', () => this.validateEmail());
        this.passwordInput.addEventListener('blur', () => this.validatePassword());
        this.confirmPasswordInput.addEventListener('blur', () => this.validateConfirmPassword());
        this.termsCheckbox.addEventListener('change', () => this.validateTerms());
        
        // Clear errors on input
        this.firstNameInput.addEventListener('input', () => this.clearError('firstName'));
        this.lastNameInput.addEventListener('input', () => this.clearError('lastName'));
        this.emailInput.addEventListener('input', () => this.clearError('email'));
        this.passwordInput.addEventListener('input', () => {
            this.clearError('password');
            if (this.confirmPasswordInput.value) {
                this.validateConfirmPassword();
            }
        });
        this.confirmPasswordInput.addEventListener('input', () => this.clearError('confirmPassword'));
        
        // Add soft press effects to inputs
        const inputs = [this.firstNameInput, this.lastNameInput, this.emailInput, 
                       this.passwordInput, this.confirmPasswordInput];
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
        
        // Validate all fields
        const isFirstNameValid = this.validateFirstName();
        const isLastNameValid = this.validateLastName();
        const isEmailValid = this.validateEmail();
        const isPasswordValid = this.validatePassword();
        const isConfirmPasswordValid = this.validateConfirmPassword();
        const isTermsValid = this.validateTerms();
        
        if (isFirstNameValid && isLastNameValid && isEmailValid && 
            isPasswordValid && isConfirmPasswordValid && isTermsValid) {
            this.submitForm();
        }
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
        const minLength = 8;
        const hasUpperCase = /[A-Z]/.test(password);
        const hasLowerCase = /[a-z]/.test(password);
        const hasNumbers = /\d/.test(password);
        const hasSpecialChar = /[!@#$%^&*(),.?":{}|<>]/.test(password);
        
        if (!password) {
            this.showError('password', 'Password is required');
            return false;
        }
        if (password.length < minLength) {
            this.showError('password', `Password must be at least ${minLength} characters long`);
            return false;
        }
        if (!hasUpperCase) {
            this.showError('password', 'Password must contain at least one uppercase letter');
            return false;
        }
        if (!hasLowerCase) {
            this.showError('password', 'Password must contain at least one lowercase letter');
            return false;
        }
        if (!hasNumbers) {
            this.showError('password', 'Password must contain at least one number');
            return false;
        }
        if (!hasSpecialChar) {
            this.showError('password', 'Password must contain at least one special character');
            return false;
        }
        
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
        const formGroup = document.getElementById(fieldName).closest('.form-group');
        const errorElement = document.getElementById(fieldName + 'Error');
        
        formGroup.classList.add('error');
        errorElement.textContent = message;
    }
    
    clearError(fieldName) {
        const formGroup = document.getElementById(fieldName).closest('.form-group');
        const errorElement = document.getElementById(fieldName + 'Error');
        
        formGroup.classList.remove('error');
        errorElement.textContent = '';
    }
    
    async submitForm() {
        // Show loading state
        this.submitButton.classList.add('loading');
        
        try {
            // Simulate API call
            await this.simulateSignup();
            
            // Hide form and show success message
            this.form.style.display = 'none';
            this.successMessage.classList.add('show');
            
            // Simulate redirect after success
            setTimeout(() => {
                // In a real app, redirect to login page or dashboard
                console.log('Redirecting to login page...');
                // window.location.href = 'login.html';
            }, 3000);
            
        } catch (error) {
            console.error('Signup error:', error);
            this.showError('email', 'An error occurred during signup. Please try again.');
        } finally {
            this.submitButton.classList.remove('loading');
        }
    }
    
    async simulateSignup() {
        // Simulate network delay
        return new Promise((resolve, reject) => {
            setTimeout(() => {
                // Simulate random success/failure for demo
                const success = Math.random() > 0.1; // 90% success rate
                if (success) {
                    resolve({ success: true, message: 'Account created successfully' });
                } else {
                    reject(new Error('Network error'));
                }
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
    new NeumorphismSignupForm();
});