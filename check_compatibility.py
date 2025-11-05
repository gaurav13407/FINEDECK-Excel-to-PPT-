"""
System Compatibility Check
Verifies all components are properly aligned and working together
"""

import os
import sys
import json

def check_tiered_converter():
    """Check tiered converter configuration"""
    print("\n" + "="*80)
    print("1. TIERED CONVERTER CONFIGURATION")
    print("="*80)
    
    try:
        sys.path.insert(0, os.path.abspath('.'))
        from src.converter.excel_to_ppt_converter import TIER_CONFIG
        
        print("\n✅ Tiered converter module loaded successfully\n")
        print(f"{'Tier':<10} {'PPT Limit':>12} {'Max Sheets':>12} {'AI Features':>15}")
        print("-" * 50)
        
        for tier, config in TIER_CONFIG.items():
            ppt_limit = "Unlimited" if config['ppt_limit'] == -1 else str(config['ppt_limit'])
            max_sheets = "Unlimited" if config['max_sheets'] == -1 else str(config['max_sheets'])
            ai_count = len(config.get('ai_features', []))
            
            print(f"{tier:<10} {ppt_limit:>12} {max_sheets:>12} {ai_count:>15}")
        
        return True
    except Exception as e:
        print(f"\n❌ Failed to load tiered converter: {str(e)}")
        return False


def check_user_model():
    """Check user model PLAN_CONFIGS"""
    print("\n" + "="*80)
    print("2. USER MODEL PLAN CONFIGS")
    print("="*80)
    
    try:
        from src.backend.app.models.user import PLAN_CONFIGS, SubscriptionPlan
        
        print("\n✅ User model loaded successfully\n")
        print(f"{'Tier':<15} {'Price':>8} {'PPT Limit':>12} {'Credits':>10} {'✓ Aligned':>10}")
        print("-" * 60)
        
        all_aligned = True
        for plan, config in PLAN_CONFIGS.items():
            ppt_limit = config.get('presentations_limit', 0)
            credit_limit = config.get('monthly_credits_limit', 0)
            price = config.get('price', 0)
            
            # Check alignment
            aligned = ppt_limit == credit_limit or (ppt_limit == -1 and credit_limit == -1)
            status = "✅" if aligned else "❌"
            
            if not aligned:
                all_aligned = False
            
            ppt_display = "Unlimited" if ppt_limit == -1 else str(ppt_limit)
            credit_display = "Unlimited" if credit_limit == -1 else str(credit_limit)
            
            plan_name = plan.value if hasattr(plan, 'value') else str(plan)
            print(f"{plan_name:<15} ${price:>7.2f} {ppt_display:>12} {credit_display:>10} {status:>10}")
        
        if all_aligned:
            print("\n✅ All plans have aligned PPT limits and credit limits!")
        else:
            print("\n⚠️  WARNING: Some plans have misaligned limits!")
        
        return all_aligned
    except Exception as e:
        print(f"\n❌ Failed to load user model: {str(e)}")
        return False


def check_file_model():
    """Check file model subscription limits"""
    print("\n" + "="*80)
    print("3. FILE MODEL SUBSCRIPTION LIMITS")
    print("="*80)
    
    try:
        from src.backend.app.models.file import get_subscription_limits
        
        print("\n✅ File model loaded successfully\n")
        
        tiers = ['free', 'basic', 'pro', 'ai_pro', 'enterprise']
        print(f"{'Tier':<15} {'File Limit':>12} {'Credits':>10} {'Storage (MB)':>15}")
        print("-" * 55)
        
        for tier in tiers:
            limits = get_subscription_limits(tier)
            file_limit = "Unlimited" if limits['monthly_file_limit'] == -1 else str(limits['monthly_file_limit'])
            credit_limit = "Unlimited" if limits['monthly_credits_limit'] == -1 else str(limits['monthly_credits_limit'])
            
            print(f"{tier:<15} {file_limit:>12} {credit_limit:>10} {limits['storage_limit_mb']:>15}")
        
        return True
    except Exception as e:
        print(f"\n❌ Failed to load file model: {str(e)}")
        return False


def check_api_endpoints():
    """Check API endpoint configurations"""
    print("\n" + "="*80)
    print("4. API ENDPOINTS CHECK")
    print("="*80)
    
    endpoints = {
        'Legacy Conversion': 'src/backend/app/api/v1/endpoints/conversions.py',
        'Tiered Conversion': 'src/backend/app/api/v1/endpoints/tiered_conversions.py',
        'User Management': 'src/backend/app/api/v1/endpoints/users.py',
    }
    
    print()
    all_exist = True
    
    for name, path in endpoints.items():
        if os.path.exists(path):
            print(f"✅ {name:<25s} → {path}")
            
            # Check for credit/PPT usage
            with open(path, 'r', encoding='utf-8') as f:
                content = f.read()
                
            uses_credits = 'deduct_user_credits' in content or 'monthly_credits' in content
            uses_ppt_count = 'this_month_conversions' in content or 'user_ppt_count' in content
            
            if uses_credits and not uses_ppt_count:
                print(f"   ⚠️  Uses ONLY credit system")
            elif uses_ppt_count and not uses_credits:
                print(f"   ✅ Uses PPT count system")
            elif uses_credits and uses_ppt_count:
                print(f"   ⚠️  Uses BOTH systems (potential conflict)")
            else:
                print(f"   ℹ️  No limit tracking found")
        else:
            print(f"❌ {name:<25s} → NOT FOUND: {path}")
            all_exist = False
    
    return all_exist


def check_converter_compatibility():
    """Check converter compatibility with tier system"""
    print("\n" + "="*80)
    print("5. CONVERTER COMPATIBILITY")
    print("="*80)
    
    try:
        from src.converter.excel_to_ppt_converter import ExcelToPPTConverter
        
        print("\n✅ Testing converter with all tiers...\n")
        
        tiers = ['free', 'basic', 'pro', 'ai_pro']
        all_success = True
        
        for tier in tiers:
            try:
                converter = ExcelToPPTConverter(tier)
                templates = converter.get_allowed_templates()
                
                template_count = len(templates) if isinstance(templates, list) else "All"
                print(f"✅ {tier.upper():<10s} → Templates: {template_count}, "
                      f"PPT Limit: {converter.config['ppt_limit']}, "
                      f"AI Features: {len(converter.config.get('ai_features', []))}")
            except Exception as e:
                print(f"❌ {tier.upper():<10s} → Failed: {str(e)}")
                all_success = False
        
        return all_success
    except Exception as e:
        print(f"\n❌ Failed to test converter: {str(e)}")
        return False


def check_ai_service():
    """Check AI service availability"""
    print("\n" + "="*80)
    print("6. AI SERVICE CHECK")
    print("="*80)
    
    try:
        from src.backend.app.services.ai_service import create_ai_service
        
        # Check for Groq API key
        groq_key = os.getenv('GROQ_API_KEY')
        if groq_key:
            print(f"\n✅ GROQ_API_KEY found: {groq_key[:20]}...")
        else:
            print("\n⚠️  GROQ_API_KEY not found in environment")
        
        # Try to create AI service
        ai_service = create_ai_service()
        print("✅ AI service created successfully")
        
        # Check available features
        features = [
            'generate_slide_title',
            'generate_slide_summary',
            'generate_data_insights',
            'recommend_template',
            'optimize_slide_layout',
            'recommend_chart_type'
        ]
        
        print("\n📋 Available AI Features:")
        for feature in features:
            if hasattr(ai_service, feature):
                print(f"   ✅ {feature}")
            else:
                print(f"   ❌ {feature} - MISSING")
        
        return True
    except Exception as e:
        print(f"\n❌ Failed to load AI service: {str(e)}")
        return False


def check_environment():
    """Check environment configuration"""
    print("\n" + "="*80)
    print("7. ENVIRONMENT CONFIGURATION")
    print("="*80)
    
    required_vars = {
        'GROQ_API_KEY': 'AI service',
        'MONGODB_URL': 'Database',
        'SECRET_KEY': 'Authentication'
    }
    
    print()
    all_present = True
    
    for var, purpose in required_vars.items():
        value = os.getenv(var)
        if value:
            # Mask sensitive values
            if 'KEY' in var or 'SECRET' in var:
                display = value[:20] + '...' if len(value) > 20 else value
            else:
                display = value
            print(f"✅ {var:<20s} → {display} ({purpose})")
        else:
            print(f"❌ {var:<20s} → NOT SET ({purpose})")
            all_present = False
    
    return all_present


def generate_summary():
    """Generate final summary"""
    print("\n" + "="*80)
    print("COMPATIBILITY CHECK SUMMARY")
    print("="*80)
    
    checks = [
        ("Tiered Converter", check_tiered_converter()),
        ("User Model", check_user_model()),
        ("File Model", check_file_model()),
        ("API Endpoints", check_api_endpoints()),
        ("Converter Compatibility", check_converter_compatibility()),
        ("AI Service", check_ai_service()),
        ("Environment", check_environment())
    ]
    
    passed = sum(1 for _, result in checks if result)
    total = len(checks)
    
    print(f"\n📊 Results: {passed}/{total} checks passed\n")
    
    for name, result in checks:
        status = "✅ PASS" if result else "❌ FAIL"
        print(f"   {status} - {name}")
    
    print("\n" + "="*80)
    
    if passed == total:
        print("🎉 ALL CHECKS PASSED! System is ready for deployment.")
    else:
        print(f"⚠️  {total - passed} checks failed. Please review and fix issues above.")
    
    print("="*80)


if __name__ == "__main__":
    print("\n🔍 Starting System Compatibility Check...")
    
    # Load environment variables
    from dotenv import load_dotenv
    load_dotenv()
    
    # Run all checks
    generate_summary()
    
    print("\n✅ Compatibility check complete!\n")
