"""
Direct Upgrade Code Sender
===========================
Simple script to generate and send upgrade codes directly to customers.
Also updates the user's subscription plan in the database.

Usage:
    python send_upgrade_code.py
"""

import asyncio
import sys
import os
from datetime import datetime, timedelta
import secrets
import string
from pathlib import Path

# Add backend path to sys.path
backend_path = Path(__file__).parent / "src" / "backend" / "app"
sys.path.insert(0, str(backend_path))

from motor.motor_asyncio import AsyncIOMotorClient
from bson import ObjectId

# Load environment variables from .env file
from dotenv import load_dotenv
load_dotenv()

# MongoDB Configuration - use your production database
MONGODB_URL = os.getenv("DATABASE_URL", "mongodb://localhost:27017")
DATABASE_NAME = os.getenv("DATABASE_NAME", "findeck_db")

# Brevo Email Configuration
BREVO_API_KEY = os.getenv("BREVO_API_KEY", "")
FROM_EMAIL = os.getenv("EMAIL_FROM", "no-reply@findeck.live")
FROM_NAME = os.getenv("MAIL_FROM_NAME", "FinDeck")


def generate_upgrade_code() -> str:
    """Generate a unique upgrade code"""
    characters = string.ascii_uppercase + string.digits
    code_suffix = ''.join(secrets.choice(characters) for _ in range(12))
    return code_suffix


def create_email_html(customer_email: str, plan: str, code: str, expires_at: datetime) -> str:
    """Create HTML email for upgrade code"""
    
    plan_features = {
        "BASIC": [
            "✓ 15 presentations per month",
            "✓ 5 template designs",
            "✓ AI-powered slide titles",
            "✓ Basic charts and graphs"
        ],
        "PRO": [
            "✓ 50 presentations per month",
            "✓ 10 premium templates",
            "✓ AI titles & summaries",
            "✓ Advanced charts with legends",
            "✓ Custom branding"
        ],
        "AI_PRO": [
            "✓ 200 presentations per month",
            "✓ All premium templates",
            "✓ Full AI suite (insights, predictions, recommendations)",
            "✓ Smart chart analyzer",
            "✓ Dark finance theme",
            "✓ Priority support",
            "✓ API access"
        ]
    }
    
    features = plan_features.get(plan, [])
    
    html = f"""
<!DOCTYPE html>
<html>
<head>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 600px;
            margin: 0 auto;
            padding: 20px;
        }}
        .container {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            border-radius: 10px;
            padding: 30px;
            color: white;
        }}
        .logo {{
            font-size: 32px;
            font-weight: bold;
            margin-bottom: 20px;
        }}
        .content {{
            background: white;
            color: #333;
            border-radius: 8px;
            padding: 30px;
            margin-top: 20px;
        }}
        .code-box {{
            background: #f8f9fa;
            border: 2px dashed #667eea;
            border-radius: 8px;
            padding: 20px;
            text-align: center;
            margin: 30px 0;
        }}
        .code {{
            font-size: 28px;
            font-weight: bold;
            color: #667eea;
            letter-spacing: 2px;
            font-family: 'Courier New', monospace;
        }}
        .features {{
            background: #f8f9fa;
            border-radius: 8px;
            padding: 20px;
            margin: 20px 0;
        }}
        .feature {{
            padding: 8px 0;
            font-size: 16px;
        }}
        .expiry {{
            color: #dc3545;
            font-weight: bold;
            margin-top: 15px;
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="logo">🎯 FinDeck</div>
        <h1 style="margin: 0;">You've Been Upgraded!</h1>
        <p style="font-size: 18px; margin-top: 10px;">Welcome to {plan.replace('_', ' ')} Plan</p>
    </div>
    
    <div class="content">
        <h2>🎉 Congratulations!</h2>
        <p>Your FinDeck account has been upgraded to <strong>{plan.replace('_', ' ')}</strong>!</p>
        
        <div class="code-box">
            <div style="font-size: 14px; color: #666; margin-bottom: 10px;">Your Upgrade Code:</div>
            <div class="code">{code}</div>
            <div class="expiry">⏰ Expires: {expires_at.strftime('%B %d, %Y')}</div>
        </div>
        
        <h3>📦 What's Included:</h3>
        <div class="features">
            {''.join(f'<div class="feature">{feature}</div>' for feature in features)}
        </div>
        
        <p style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #eee; color: #666; font-size: 14px;">
            This code is valid for 30 days. Your subscription has been automatically activated!
        </p>
        <p style="color: #999; font-size: 12px; text-align: center; margin-top: 20px;">
            © 2025 FinDeck. All rights reserved.
        </p>
    </div>
</body>
</html>
"""
    return html


async def send_email_brevo(to_email: str, subject: str, html_content: str) -> bool:
    """Send email using Brevo (Sendinblue) API"""
    try:
        import aiohttp
        
        url = "https://api.brevo.com/v3/smtp/email"
        
        headers = {
            "accept": "application/json",
            "api-key": BREVO_API_KEY,
            "content-type": "application/json"
        }
        
        payload = {
            "sender": {
                "name": FROM_NAME,
                "email": FROM_EMAIL
            },
            "to": [
                {
                    "email": to_email,
                    "name": to_email.split('@')[0]
                }
            ],
            "subject": subject,
            "htmlContent": html_content
        }
        
        async with aiohttp.ClientSession() as session:
            async with session.post(url, json=payload, headers=headers) as response:
                if response.status == 201:
                    result = await response.json()
                    print(f"✅ Email sent to {to_email}")
                    print(f"   Message ID: {result.get('messageId', 'N/A')}")
                    return True
                else:
                    error_text = await response.text()
                    print(f"❌ Failed to send email: HTTP {response.status}")
                    print(f"   Error: {error_text}")
                    return False
                    
    except Exception as e:
        print(f"❌ Failed to send email: {str(e)}")
        print(f"   You can manually send the code to the customer")
        return False


async def update_user_subscription(email: str, plan: str, code: str):
    """Update user's subscription in database"""
    
    print(f"🔌 Connecting to database...")
    print(f"   URL: {MONGODB_URL[:50]}...")
    print(f"   Database: {DATABASE_NAME}")
    
    # Connect to MongoDB
    client = AsyncIOMotorClient(MONGODB_URL)
    db = client[DATABASE_NAME]
    users_collection = db["users"]
    
    try:
        # Find user
        user = await users_collection.find_one({"email": email})
        
        if not user:
            print(f"❌ User not found: {email}")
            return False
        
        print(f"✅ Found user: {email}")
        print(f"   Current plan: {user.get('subscription', {}).get('plan', 'FREE')}")
        
        # Calculate new expiry date (1 year from now)
        now = datetime.utcnow()
        subscription_ends_at = now + timedelta(days=365)
        
        # Define plan limits (aligned with backend)
        plan_limits = {
            "FREE": {"credits": 1, "presentations": 1, "price": 0.0},
            "BASIC": {"credits": 7, "presentations": 7, "price": 25.0},
            "PRO": {"credits": 15, "presentations": 15, "price": 49.0},
            "AI_PRO": {"credits": 200, "presentations": 200, "price": 99.0},  # FIXED: Use 200 instead of -1 for unlimited
            "ENTERPRISE": {"credits": 1000, "presentations": 1000, "price": 99.99}
        }
        
        # Get limits for the plan
        limits = plan_limits.get(plan, plan_limits["FREE"])
        
        # Update subscription with all required fields
        update_result = await users_collection.update_one(
            {"email": email},
            {
                "$set": {
                    "subscription.plan": plan.lower(),  # FIXED: Convert to lowercase to match enum
                    "subscription.status": "active",
                    "subscription.starts_at": now,
                    "subscription.ends_at": subscription_ends_at,
                    "subscription.upgraded_at": now,
                    "subscription.upgrade_method": "manual_code",
                    "subscription.upgrade_code": code,
                    "subscription.monthly_credits_limit": limits["credits"],  # FIXED: Set proper credit limit
                    "subscription.monthly_credits_used": 0,  # Reset credits used
                    "subscription.presentations_limit": limits["presentations"],  # FIXED: Set proper presentation limit
                    "subscription.price_per_month": limits["price"],  # FIXED: Set proper price
                    "subscription.credits_remaining": limits["credits"],  # FIXED: Set remaining credits
                    "subscription.credits_reset_date": now,
                    "updated_at": now
                }
            }
        )
        
        if update_result.modified_count > 0:
            print(f"✅ Subscription updated successfully!")
            print(f"   New plan: {plan}")
            print(f"   Expires: {subscription_ends_at.strftime('%Y-%m-%d')}")
            return True
        else:
            print(f"⚠️  No changes made to subscription")
            return False
            
    except Exception as e:
        print(f"❌ Database error: {str(e)}")
        return False
    finally:
        client.close()


async def save_upgrade_code(email: str, plan: str, code: str):
    """Save upgrade code to database"""
    
    print(f"🔌 Connecting to database for code storage...")
    
    client = AsyncIOMotorClient(MONGODB_URL)
    db = client[DATABASE_NAME]
    codes_collection = db["upgrade_codes"]
    
    try:
        now = datetime.utcnow()
        expires_at = now + timedelta(days=30)
        
        code_document = {
            "code": code,
            "plan": plan,
            "customer_email": email,
            "generated_by": "system",
            "generated_by_email": "admin@findeck.com",
            "generated_at": now,
            "expires_at": expires_at,
            "is_redeemed": True,  # Already applied
            "redeemed_at": now,
            "redeemed_by_email": email,
            "duration_months": 12,
            "notes": "Direct upgrade code - subscription already activated",
            "meta": {}
        }
        
        await codes_collection.insert_one(code_document)
        print(f"✅ Code saved to database")
        return True
        
    except Exception as e:
        print(f"❌ Failed to save code: {str(e)}")
        return False
    finally:
        client.close()


async def send_upgrade_code(email: str, plan: str):
    """Main function to send upgrade code and update subscription"""
    
    print("="*80)
    print("🎁 SENDING UPGRADE CODE")
    print("="*80)
    print(f"\n📧 Customer Email: {email}")
    print(f"📦 Plan: {plan}")
    print()
    
    # Generate code
    code_suffix = generate_upgrade_code()
    full_code = f"{plan}-{code_suffix}"
    
    print(f"🔑 Generated Code: {full_code}")
    print()
    
    # Update subscription in database
    print("📊 Updating subscription in database...")
    db_updated = await update_user_subscription(email, plan, full_code)
    
    if not db_updated:
        print("\n❌ Failed to update subscription. Aborting.")
        return
    
    # Save code to database
    print("\n💾 Saving code record...")
    await save_upgrade_code(email, plan, full_code)
    
    # Send email
    print("\n📧 Sending email...")
    expires_at = datetime.utcnow() + timedelta(days=30)
    
    email_html = create_email_html(email, plan, full_code, expires_at)
    
    # Try to send email using Brevo
    await send_email_brevo(
        to_email=email,
        subject=f"🎉 Your FinDeck {plan.replace('_', ' ')} Upgrade",
        html_content=email_html
    )
    
    print("\n" + "="*80)
    print("✅ UPGRADE COMPLETE!")
    print("="*80)
    print(f"\n✨ Summary:")
    print(f"   • Email: {email}")
    print(f"   • Plan: {plan}")
    print(f"   • Code: {full_code}")
    print(f"   • Expires: {expires_at.strftime('%B %d, %Y')}")
    print(f"   • Subscription: Active until {expires_at.strftime('%Y-%m-%d')}")
    print()


async def main():
    """Main entry point"""
    
    print("\n" + "="*80)
    print("🎯 FINDECK UPGRADE CODE SENDER")
    print("="*80)
    print()
    
    # Configuration for your email
    customer_email = "gaurav13407@outlook.com"
    plan = "AI_PRO"
    
    print(f"Target Configuration:")
    print(f"  Email: {customer_email}")
    print(f"  Plan:  {plan}")
    print()
    
    confirm = input("⚠️  This will update the database. Continue? (yes/no): ")
    
    if confirm.lower() not in ['yes', 'y']:
        print("\n❌ Cancelled by user")
        return
    
    print()
    await send_upgrade_code(customer_email, plan)
    
    print("\n💡 Next Steps:")
    print("   1. Check your email for the upgrade notification")
    print("   2. Log into your account to see the new plan")
    print("   3. Start using AI Pro features!")
    print()


if __name__ == "__main__":
    # Check if aiohttp is installed
    try:
        import aiohttp
    except ImportError:
        print("⚠️  Installing aiohttp for Brevo email support...")
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", "aiohttp"])
        print("✅ aiohttp installed")
    
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\n\n⚠️  Interrupted by user")
    except Exception as e:
        print(f"\n\n❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()
