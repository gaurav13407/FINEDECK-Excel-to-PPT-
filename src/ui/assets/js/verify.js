// verify.js - handles verification UI interactions
(function(){
    function qs(name) {
        return new URLSearchParams(window.location.search).get(name) || '';
    }

    const email = qs('email');
    const purpose = qs('purpose') || 'email_verification';
    const plan = qs('plan') || '';

    const emailText = document.getElementById('emailText');
    const purposeText = document.getElementById('purposeText');
    const codeInput = document.getElementById('codeInput');
    const verifyBtn = document.getElementById('verifyBtn');
    const resendBtn = document.getElementById('resendBtn');
    const cancelBtn = document.getElementById('cancelBtn');
    const messageDiv = document.getElementById('message');

    if (emailText) emailText.textContent = email || 'your email';
    if (purposeText) {
        const label = purpose === 'login_verification' ? 'Login verification' : purpose === 'redeem' ? 'Redeem code' : 'Email verification';
        purposeText.textContent = label + (plan ? ` — ${plan}` : '');
    }

    // Show code format info to the user
    const codeInfo = 'Code format: 8 characters (A-Z and digits 2-9), displayed as XXXX-XXXX (9 chars including the hyphen).';
    showMessage(codeInfo, 'info');

    function showMessage(text, cls){
        messageDiv.className = cls || '';
        messageDiv.textContent = text;
    }

    async function apiPost(path, body){
        try {
            const res = await fetch(path, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(body)
            });

            const text = await res.text();
            let data = null;
            try { data = text ? JSON.parse(text) : null; } catch(e) { data = { raw: text }; }

            if (!res.ok) {
                // Attach status and server message when available
                const err = new Error(data?.detail || data?.message || res.statusText || 'Request failed');
                err.status = res.status;
                err.body = data;
                throw err;
            }

            return data;
        } catch (err) {
            console.warn('API request failed', err);
            throw err; // let caller handle
        }
    }

    async function sendCode() {
        // disable resend/verify while sending
        if (resendBtn) resendBtn.disabled = true;
        if (verifyBtn) verifyBtn.disabled = true;
        showMessage('Sending code…');

        // If email not provided in query params, prompt the user
        if (!email) {
            const entered = window.prompt('Enter your email address to receive the verification code:');
            if (!entered) {
                showMessage('Email required to send verification code.', 'error');
                if (resendBtn) resendBtn.disabled = false;
                if (verifyBtn) verifyBtn.disabled = false;
                return;
            }
            email = entered.trim();
            if (emailText) emailText.textContent = email;
        }

        const payload = { email: email, purpose: purpose, plan: plan };
        try {
            // Determine endpoint from global apiConfig when available
            const base = window.apiConfig ? window.apiConfig.API_BASE : '';
            const url = base ? `${base}/codes/send` : '/api/v1/codes/send';
            const result = await apiPost(url, payload);
            if (result && (result.ok || result.success)) {
                showMessage('Verification code sent. Check your inbox (or spam).', 'success');
            } else {
                showMessage('Code request completed. Check your email.', 'success');
            }
        } catch (err) {
            // 401 / 403 -> prompt login
            if (err.status === 401) {
                showMessage('You must be logged in to request this code. Please sign in.', 'error');
            } else {
                const serverMsg = err.body?.detail || err.body?.message || err.message;
                showMessage('Failed to send code: ' + serverMsg, 'error');
            }
        } finally {
            if (resendBtn) resendBtn.disabled = false;
            if (verifyBtn) verifyBtn.disabled = false;
        }
    }

    async function verifyCode() {
        const raw = codeInput.value.trim();
        if (!raw) { showMessage('Enter the verification code', 'error'); return; }
        showMessage('Verifying…');
        // disable buttons during verify
        if (verifyBtn) verifyBtn.disabled = true;
        if (resendBtn) resendBtn.disabled = true;

            const payload = { code: raw, email: email, purpose: purpose };
        try {
            const base = window.apiConfig ? window.apiConfig.API_BASE : '';
            const url = base ? `${base}/codes/verify` : '/api/v1/codes/verify';
            const result = await apiPost(url, payload);
            if (result && (result.ok || result.success)) {
                showMessage('Verified successfully! Redirecting…', 'success');
                setTimeout(() => {
                    // Redirect based on purpose
                    if (purpose === 'redeem') window.location.href = 'mainpage.html';
                    else window.location.href = 'mainpage.html';
                }, 900);
            } else {
                showMessage('Verification completed. Redirecting…', 'success');
                setTimeout(()=> window.location.href = 'mainpage.html', 900);
            }
        } catch (err) {
            if (err.status === 400) {
                const serverMsg = err.body?.detail || err.body?.message || 'Invalid or expired code';
                showMessage(serverMsg, 'error');
            } else if (err.status === 401) {
                showMessage('Authentication required. Please sign in to verify this code.', 'error');
            } else {
                const serverMsg = err.body?.detail || err.body?.message || err.message;
                showMessage('Verification failed: ' + serverMsg, 'error');
            }
        } finally {
            if (verifyBtn) verifyBtn.disabled = false;
            if (resendBtn) resendBtn.disabled = false;
        }
    }

    // Format code input: uppercase, allow A-Z and digits 2-9, auto-insert hyphen after 4 characters
    function formatCodeInput(value) {
        if (!value) return '';
        // remove non-alphanumeric and hyphen
        let cleaned = value.toUpperCase().replace(/[^A-Z0-9]/g, '');
        // keep only characters from allowed alphabet (A-Z and 2-9)
        cleaned = cleaned.split('').filter(ch => {
            if (/[A-Z]/.test(ch)) return true;
            if (/[2-9]/.test(ch)) return true;
            return false;
        }).join('');
        // limit to 8 chars (without hyphen)
        cleaned = cleaned.slice(0, 8);
        // insert hyphen after 4 chars if length >4
        if (cleaned.length > 4) {
            return cleaned.slice(0,4) + '-' + cleaned.slice(4);
        }
        return cleaned;
    }

    function handleInputFormat(e){
        const el = e.target;
        const formatted = formatCodeInput(el.value);
        el.value = formatted;
    }

    if (codeInput) {
        // Ensure maxlength matches pattern (9 including hyphen)
        codeInput.setAttribute('maxlength', '9');
        codeInput.addEventListener('input', handleInputFormat);
        codeInput.addEventListener('paste', (ev)=>{
            setTimeout(()=> handleInputFormat({ target: codeInput }), 0);
        });
    }

    // wire events
    if (resendBtn) resendBtn.addEventListener('click', (e)=>{ e.preventDefault(); sendCode(); });
    if (verifyBtn) verifyBtn.addEventListener('click', (e)=>{ e.preventDefault(); verifyCode(); });
    if (cancelBtn) cancelBtn.addEventListener('click', (e)=>{ e.preventDefault(); window.location.href = 'mainpage.html'; });

    // auto-send code when page opens (if email provided)
    if (email) {
        // slight delay for UX
        setTimeout(() => sendCode(), 450);
    }

})();
