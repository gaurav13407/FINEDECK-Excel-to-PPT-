# 🎁 Plan Upgrade Code System

A complete system for generating and managing subscription upgrade codes for FinDeck customers.

## 📋 Overview

This system allows admins to:
- Generate unique upgrade codes for any subscription tier
- Automatically send upgrade codes via email to customers
- Track all codes (redeemed, active, expired)
- Set custom durations (1 month, 3 months, 6 months, 12 months, lifetime)

Customers can:
- Receive upgrade codes via professional email
- View all codes sent to their email
- Redeem codes to instantly upgrade their plan
- See all features unlocked after upgrade

## 🚀 Quick Start

### 1. Start Backend Server

```bash
cd src/backend/app
uvicorn main:app --reload
```

### 2. Access Admin Interface

Open `src/ui/admin-upgrade-codes.html` in your browser or visit:
```
http://localhost:8000/admin-upgrade-codes.html
```

### 3. Generate Upgrade Code

1. Enter customer email
2. Select plan (BASIC, PRO, or AI_PRO)
3. Choose duration
4. Add optional notes
5. Click "Generate Upgrade Code & Send Email"

The customer will receive a professional email with:
- ✅ Unique upgrade code
- ✅ Plan features list
- ✅ Activation instructions
- ✅ Expiration date

## 🔌 API Endpoints

### Admin Endpoints

#### Generate Upgrade Code
```http
POST /api/v1/upgrades/generate-upgrade-code
Authorization: Bearer {admin_token}
Content-Type: application/json

{
  "email": "customer@example.com",
  "plan": "BASIC",
  "duration_months": 1,
  "notes": "Promotional upgrade"
}
```

**Response:**
```json
{
  "success": true,
  "message": "Upgrade code generated and sent to customer@example.com",
  "code": "BASIC-A1B2C3D4E5F6",
  "plan": "BASIC",
  "expires_at": "2025-12-05T10:30:00",
  "customer_email": "customer@example.com",
  "code_id": "507f1f77bcf86cd799439011"
}
```

#### View All Codes
```http
GET /api/v1/upgrades/admin/all-upgrade-codes?include_redeemed=true&include_expired=false
Authorization: Bearer {admin_token}
```

### Customer Endpoints

#### View My Codes
```http
GET /api/v1/upgrades/my-upgrade-codes
Authorization: Bearer {customer_token}
```

**Response:**
```json
{
  "codes": [
    {
      "code": "BASIC-A1B2C3D4E5F6",
      "plan": "BASIC",
      "generated_at": "2025-11-05T10:30:00",
      "expires_at": "2025-12-05T10:30:00",
      "is_redeemed": false,
      "redeemed_at": null,
      "is_expired": false,
      "can_redeem": true
    }
  ]
}
```

#### Redeem Upgrade Code
```http
POST /api/v1/upgrades/redeem-upgrade-code
Authorization: Bearer {customer_token}
Content-Type: application/json

{
  "code": "BASIC-A1B2C3D4E5F6"
}
```

**Response:**
```json
{
  "success": true,
  "message": "Congratulations! Your account has been upgraded to Basic",
  "new_plan": "BASIC",
  "subscription_ends_at": "2025-12-05T10:30:00",
  "features_unlocked": [
    "15 presentations per month",
    "5 template designs",
    "AI-powered slide titles",
    "Basic charts and graphs"
  ]
}
```

## 📧 Email Template

Customers receive a beautifully designed email with:

```
🎯 FinDeck
You've Been Upgraded!
Welcome to Basic Plan

🎉 Congratulations!
Your FinDeck account has been upgraded to Basic Plan!

┌─────────────────────────┐
│ Your Upgrade Code:      │
│ BASIC-A1B2C3D4E5F6     │
│ ⏰ Expires: Dec 5, 2025 │
└─────────────────────────┘

📦 What's Included:
✓ 15 presentations per month
✓ 5 template designs
✓ AI-powered slide titles
✓ Basic charts and graphs

🚀 How to Activate:
1. Log in to your FinDeck account
2. Go to Account Settings → Subscription
3. Click "Redeem Upgrade Code"
4. Enter the code above
5. Start creating amazing presentations!

[Activate Now →]
```

## 🔐 Security Features

✅ **Code Validation**
- Unique codes with plan prefix (e.g., `BASIC-A1B2C3D4E5F6`)
- 30-day expiration by default
- One-time use only

✅ **Email Verification**
- Codes are tied to specific email addresses
- Prevents unauthorized redemption

✅ **Audit Trail**
- Track who generated each code
- Record redemption timestamp
- Store internal notes

✅ **Access Control**
- Admin-only code generation
- Customer can only redeem their own codes
- Role-based permissions

## 🧪 Testing

### Run Test Script

```bash
python test_upgrade_codes.py
```

This will:
1. Login as admin
2. Generate upgrade code for test customer
3. Login as customer
4. Check customer's codes
5. Redeem the code
6. Verify plan upgrade
7. Show admin dashboard

### Manual Testing

1. **Create Admin Account** (if needed)
```bash
python src/backend/app/scripts/create_admin.py
```

2. **Create Test Customer** (if needed)
```bash
python src/backend/app/scripts/create_test_user.py
```

3. **Generate Code via API**
```bash
curl -X POST http://localhost:8000/api/v1/upgrades/generate-upgrade-code \
  -H "Authorization: Bearer YOUR_ADMIN_TOKEN" \
  -H "Content-Type: application/json" \
  -d '{
    "email": "customer@example.com",
    "plan": "BASIC",
    "duration_months": 1
  }'
```

4. **Redeem Code**
```bash
curl -X POST http://localhost:8000/api/v1/upgrades/redeem-upgrade-code \
  -H "Authorization: Bearer YOUR_CUSTOMER_TOKEN" \
  -H "Content-Type: application/json" \
  -d '{"code": "BASIC-A1B2C3D4E5F6"}'
```

## 📊 Database Schema

### Upgrade Codes Collection

```javascript
{
  "code": "BASIC-A1B2C3D4E5F6",           // Unique code
  "plan": "BASIC",                          // Target plan
  "customer_email": "customer@example.com", // Recipient
  "generated_by": "ObjectId(...)",          // Admin user ID
  "generated_by_email": "admin@findeck.com",
  "generated_at": ISODate("2025-11-05"),
  "expires_at": ISODate("2025-12-05"),
  "is_redeemed": false,
  "redeemed_at": null,
  "redeemed_by": null,
  "redeemed_by_email": null,
  "duration_months": 1,
  "notes": "Promotional upgrade",
  "meta": {
    "ip": "192.168.1.1",
    "user_agent": "Mozilla/5.0..."
  }
}
```

## 🎯 Use Cases

### 1. Customer Support
- Customer has billing issue → Send free month upgrade code
- Customer provides feedback → Reward with PRO upgrade

### 2. Marketing Campaigns
- Black Friday promotion → Generate bulk codes
- Referral rewards → Send upgrade codes to referrers

### 3. Partnerships
- Partner companies → Bulk codes for employees
- Event sponsors → Premium codes for attendees

### 4. Customer Retention
- User about to churn → Offer upgrade code
- Long-term user → Loyalty upgrade reward

## 🔄 Integration with Frontend

Add redemption interface in user settings:

```html
<!-- In user dashboard/settings -->
<div class="upgrade-section">
    <h3>Have an Upgrade Code?</h3>
    <input type="text" id="upgradeCode" placeholder="Enter code (e.g., BASIC-ABC123)">
    <button onclick="redeemCode()">Redeem Code</button>
</div>

<script>
async function redeemCode() {
    const code = document.getElementById('upgradeCode').value;
    const response = await fetch('/api/v1/upgrades/redeem-upgrade-code', {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ code })
    });
    
    if (response.ok) {
        const data = await response.json();
        alert(`Success! Upgraded to ${data.new_plan}`);
        location.reload();
    }
}
</script>
```

## 📈 Monitoring & Analytics

Track key metrics:
- ✅ Codes generated per day/week/month
- ✅ Redemption rate (redeemed / total generated)
- ✅ Average time to redemption
- ✅ Most popular upgrade tier
- ✅ Expired codes (wasted opportunities)

## 🛠️ Customization

### Change Code Format
Edit `generate_upgrade_code()` in `plan_upgrades.py`:
```python
def generate_upgrade_code() -> str:
    # Custom format: PLAN-XXXX-XXXX-XXXX
    return f"{secrets.token_hex(6).upper()}"
```

### Modify Email Template
Edit `create_upgrade_email_html()` function to match your brand.

### Add More Plans
Add to `valid_plans` list and feature descriptions.

## 🚨 Troubleshooting

### Email Not Sending
- Check `settings.EMAIL_*` configuration
- Verify SMTP credentials
- Check email service logs

### Code Not Working
- Verify code hasn't been redeemed
- Check expiration date
- Ensure correct email address

### Permission Denied
- Verify user has admin role
- Check authentication token
- Review role-based access settings

## 📞 Support

For issues or questions:
- 📧 Email: support@findeck.com
- 📚 Docs: https://docs.findeck.com
- 💬 Discord: https://discord.gg/findeck

---

**Built with ❤️ for FinDeck customers**
