"""
Test Excel to PPT Converter with All Tiers
Demonstrates Free, Basic, Pro, and AI Pro functionality
"""

import os
import sys
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

from src.converter.excel_to_ppt_converter import convert_excel_to_ppt, TIER_CONFIG

def test_all_tiers():
    """Test converter with all subscription tiers"""
    
    # Use sample Excel file
    excel_file = "examples/Sample_pnl.xlsx"
    
    if not os.path.exists(excel_file):
        print(f"❌ Sample file not found: {excel_file}")
        print("Please make sure you have sample Excel files in the examples/ folder")
        return
    
    print("=" * 80)
    print("TESTING EXCEL TO PPT CONVERTER - ALL TIERS")
    print("=" * 80)
    
    # Test each tier
    tiers = ['free', 'basic', 'pro', 'ai_pro']
    
    for tier in tiers:
        print(f"\n{'='*80}")
        print(f"TESTING: {TIER_CONFIG[tier]['name']}")
        print(f"{'='*80}")
        
        print(f"\nTier Configuration:")
        print(f"  - PPT Limit: {TIER_CONFIG[tier]['ppt_limit'] if TIER_CONFIG[tier]['ppt_limit'] > 0 else 'Unlimited'}")
        print(f"  - Templates: {TIER_CONFIG[tier]['templates']}")
        print(f"  - AI Features: {', '.join(TIER_CONFIG[tier]['ai_features']) if TIER_CONFIG[tier]['ai_features'] else 'None'}")
        print(f"  - Multi-sheet: {TIER_CONFIG[tier]['multi_sheet']}")
        
        # Output file
        output_file = f"examples/demo_PPT/test_{tier}_output.pptx"
        
        # Convert
        print(f"\nConverting...")
        result = convert_excel_to_ppt(
            excel_path=excel_file,
            output_path=output_file,
            user_tier=tier,
            user_ppt_count=0  # First PPT of the month
        )
        
        # Show results
        if result['success']:
            print(f"\n✅ SUCCESS!")
            print(f"  - Output: {result['output_path']}")
            print(f"  - Slides created: {result['slides_created']}")
            print(f"  - Template used: {result['template_used']}")
            print(f"  - AI features used: {', '.join(result['ai_features_used']) if result['ai_features_used'] else 'None'}")
            
            if 'ai_usage' in result and result['ai_usage']['tokens'] > 0:
                print(f"\n  📊 AI Usage:")
                print(f"     - Tokens: {result['ai_usage']['tokens']}")
                print(f"     - Cost: ${result['ai_usage']['cost']:.6f}")
        else:
            print(f"\n❌ FAILED: {result.get('error', 'Unknown error')}")
    
    print(f"\n{'='*80}")
    print("ALL TESTS COMPLETE!")
    print(f"{'='*80}")
    print("\nCheck the examples/demo_PPT/ folder for generated presentations:")
    print("  - test_free_output.pptx (Free tier)")
    print("  - test_basic_output.pptx (Basic tier with AI titles)")
    print("  - test_pro_output.pptx (Pro tier with AI titles + template selection)")
    print("  - test_ai_pro_output.pptx (AI Pro tier with all AI features)")

def test_limit_enforcement():
    """Test PPT limit enforcement"""
    
    print("\n" + "=" * 80)
    print("TESTING PPT LIMIT ENFORCEMENT")
    print("=" * 80)
    
    excel_file = "examples/Sample_pnl.xlsx"
    
    # Test Free tier limit (1 PPT)
    print("\n1. Testing Free tier (limit: 1 PPT/month)")
    print("   Attempting to create 2nd PPT...")
    
    result = convert_excel_to_ppt(
        excel_path=excel_file,
        output_path="examples/demo_PPT/test_free_2nd.pptx",
        user_tier='free',
        user_ppt_count=1  # Already created 1
    )
    
    if not result['success'] and result.get('upgrade_required'):
        print(f"   ✅ Correctly blocked: {result['error']}")
    else:
        print(f"   ❌ Should have been blocked!")
    
    # Test Basic tier limit (7 PPTs)
    print("\n2. Testing Basic tier (limit: 7 PPTs/month)")
    print("   Attempting to create 8th PPT...")
    
    result = convert_excel_to_ppt(
        excel_path=excel_file,
        output_path="examples/demo_PPT/test_basic_8th.pptx",
        user_tier='basic',
        user_ppt_count=7  # Already created 7
    )
    
    if not result['success'] and result.get('upgrade_required'):
        print(f"   ✅ Correctly blocked: {result['error']}")
    else:
        print(f"   ❌ Should have been blocked!")
    
    # Test Pro tier limit (15 PPTs)
    print("\n3. Testing Pro tier (limit: 15 PPTs/month)")
    print("   Attempting to create 16th PPT...")
    
    result = convert_excel_to_ppt(
        excel_path=excel_file,
        output_path="examples/demo_PPT/test_pro_16th.pptx",
        user_tier='pro',
        user_ppt_count=15  # Already created 15
    )
    
    if not result['success'] and result.get('upgrade_required'):
        print(f"   ✅ Correctly blocked: {result['error']}")
    else:
        print(f"   ❌ Should have been blocked!")
    
    # Test AI Pro (unlimited)
    print("\n4. Testing AI Pro tier (unlimited)")
    print("   Creating 101st PPT...")
    
    result = convert_excel_to_ppt(
        excel_path=excel_file,
        output_path="examples/demo_PPT/test_ai_pro_101st.pptx",
        user_tier='ai_pro',
        user_ppt_count=100  # Already created 100
    )
    
    if result['success']:
        print(f"   ✅ Correctly allowed (unlimited tier)")
    else:
        print(f"   ❌ Should have been allowed!")

def compare_tiers():
    """Show tier comparison"""
    
    print("\n" + "=" * 80)
    print("TIER COMPARISON")
    print("=" * 80)
    
    print(f"\n{'Feature':<30} {'Free':<15} {'Basic':<15} {'Pro':<15} {'AI Pro':<15}")
    print("-" * 90)
    print(f"{'Price':<30} {'$0':<15} {'$25/month':<15} {'$49/month':<15} {'$99/month':<15}")
    print(f"{'PPT Limit':<30} {'1/month':<15} {'7/month':<15} {'15/month':<15} {'Unlimited':<15}")
    print(f"{'Templates':<30} {'1 basic':<15} {'1 basic':<15} {'10 pro':<15} {'10 pro':<15}")
    print(f"{'Multi-sheet':<30} {'No':<15} {'Yes (5)':<15} {'Yes (20)':<15} {'Unlimited':<15}")
    print(f"{'AI Titles':<30} {'No':<15} {'Yes':<15} {'Yes':<15} {'Yes':<15}")
    print(f"{'AI Template Select':<30} {'No':<15} {'No':<15} {'Yes':<15} {'Yes':<15}")
    print(f"{'AI Summaries':<30} {'No':<15} {'No':<15} {'No':<15} {'Yes':<15}")
    print(f"{'AI Insights':<30} {'No':<15} {'No':<15} {'No':<15} {'Yes':<15}")
    print(f"{'AI Layout Optimize':<30} {'No':<15} {'No':<15} {'No':<15} {'Yes':<15}")
    print(f"{'AI Chart Recommend':<30} {'No':<15} {'No':<15} {'No':<15} {'Yes':<15}")
    print("-" * 90)
    print(f"{'Cost per PPT':<30} {'$0':<15} {'~$0.0001':<15} {'~$0.0003':<15} {'~$0.001':<15}")
    print(f"{'Profit Margin':<30} {'-':<15} {'99.9%':<15} {'99.9%':<15} {'99.9%':<15}")

if __name__ == "__main__":
    # Show tier comparison first
    compare_tiers()
    
    # Test all tiers
    test_all_tiers()
    
    # Test limit enforcement
    test_limit_enforcement()
    
    print("\n" + "=" * 80)
    print("🎉 ALL TESTS COMPLETE!")
    print("=" * 80)
