// Quick script to check localStorage data
// Run this in browser console or use the storage-inspector.html page

console.log('='.repeat(60));
console.log('📦 FINDECK LOCALSTORAGE DATA INSPECTOR');
console.log('='.repeat(60));

// 1. Check finDeckAuth
console.log('\n1️⃣ finDeckAuth (Login Data):');
const authData = localStorage.getItem('finDeckAuth');
if (authData) {
    try {
        const parsed = JSON.parse(authData);
        console.log('✅ FOUND!');
        console.log(JSON.stringify(parsed, null, 2));
        if (parsed.user && parsed.user.plan) {
            console.log(`\n🎯 PLAN DETECTED: ${parsed.user.plan}`);
        }
    } catch (e) {
        console.log('❌ Invalid JSON:', authData);
    }
} else {
    console.log('❌ NOT FOUND');
}

// 2. Check authToken
console.log('\n2️⃣ authToken (API Token):');
const token = localStorage.getItem('authToken') || localStorage.getItem('token');
if (token) {
    console.log('✅ FOUND!');
    console.log(`Token (first 50 chars): ${token.substring(0, 50)}...`);
} else {
    console.log('❌ NOT FOUND');
}

// 3. Check user
console.log('\n3️⃣ user (Cached User Data):');
const userData = localStorage.getItem('user');
if (userData) {
    try {
        const parsed = JSON.parse(userData);
        console.log('✅ FOUND!');
        console.log(JSON.stringify(parsed, null, 2));
        
        console.log('\n🔍 Checking Plan Fields:');
        const fields = ['ai_pro', 'subscription_tier', 'tier', 'plan', 'subscription_plan', 'planType'];
        fields.forEach(field => {
            const value = parsed[field];
            if (value !== undefined) {
                console.log(`  ✅ ${field}: ${value}`);
            } else {
                console.log(`  ⚪ ${field}: undefined`);
            }
        });
    } catch (e) {
        console.log('❌ Invalid JSON:', userData);
    }
} else {
    console.log('❌ NOT FOUND');
}

// 4. All keys
console.log('\n4️⃣ All LocalStorage Keys:');
console.log(`Total keys: ${localStorage.length}`);
for (let i = 0; i < localStorage.length; i++) {
    const key = localStorage.key(i);
    const value = localStorage.getItem(key);
    console.log(`  • ${key} (${value.length} chars)`);
}

// 5. Diagnosis
console.log('\n' + '='.repeat(60));
console.log('🔍 DIAGNOSIS:');
console.log('='.repeat(60));

if (!authData && !userData && !token) {
    console.log('❌ NOT LOGGED IN - No authentication data found!');
    console.log('👉 Please log in at login.html');
} else {
    if (authData) {
        try {
            const parsed = JSON.parse(authData);
            if (parsed.user && parsed.user.plan) {
                console.log(`✅ LOGGED IN - Plan: ${parsed.user.plan}`);
            } else {
                console.log('⚠️ Logged in but no plan found in finDeckAuth');
            }
        } catch (e) {}
    }
    
    if (token) {
        console.log('✅ API token exists - Backend calls should work');
    } else {
        console.log('⚠️ No API token - Backend calls will fail');
    }
    
    if (userData) {
        try {
            const parsed = JSON.parse(userData);
            const fields = ['ai_pro', 'subscription_tier', 'tier', 'plan', 'subscription_plan', 'planType'];
            const foundField = fields.find(f => parsed[f] !== undefined);
            if (foundField) {
                console.log(`✅ Plan field found: ${foundField} = ${parsed[foundField]}`);
            } else {
                console.log('❌ User object exists but NO PLAN FIELD FOUND!');
            }
        } catch (e) {}
    }
}

console.log('\n' + '='.repeat(60));
console.log('📋 INSTRUCTIONS:');
console.log('='.repeat(60));
console.log('1. Copy all the output above');
console.log('2. Share it with me');
console.log('3. If plan is found but dashboard shows FREE:');
console.log('   - Clear cache (Ctrl+Shift+Delete)');
console.log('   - Or run: localStorage.clear()');
console.log('   - Then log in again');
console.log('='.repeat(60));
