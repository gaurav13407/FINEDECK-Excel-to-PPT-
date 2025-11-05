"""
Generate PPTs for ALL TIERS + GOOGL Financial Analysis
- BASIC tier (7 slides, AI summary only)
- PRO tier (7 slides, AI summary + insights)
- AI_PRO tier (8 slides, full AI features)
- GOOGL Financial Deep Dive
"""

import sys
import os
from pathlib import Path

project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from src.converter.excel_to_ppt_converter import ExcelToPPTConverter

print("=" * 100)
print("🎯 GENERATING ALL TIERS + GOOGL FINANCIAL ANALYSIS")
print("=" * 100)

excel_file = "examples/Company_Data/company_bundle.xlsx"
output_dir = "examples/professional_demo"
os.makedirs(output_dir, exist_ok=True)

user_metadata = {
    'name': 'Financial Analysis Team',
    'company': 'Global Investment Partners',
    'email': 'analytics@globalinvest.com'
}

# ============================================================================
# TIER 1: BASIC (AI Summary only)
# ============================================================================
print("\n" + "=" * 100)
print("📊 TIER 1: BASIC (7 slides, AI executive summary)")
print("=" * 100)

converter_basic = ExcelToPPTConverter(
    user_tier='basic',
    user_id='tier_basic',
    user_metadata=user_metadata
)

output_basic = os.path.join(output_dir, "Tech_Stocks_BASIC_Tier.pptx")

result_basic = converter_basic.convert_professional(
    excel_path=excel_file,
    output_path=output_basic,
    presentation_title="Tech Stocks Q4 2025 - Basic Analysis",
    user_ppt_count=0
)

if result_basic['success']:
    file_size = os.path.getsize(output_basic) / 1024
    print(f"\n✅ BASIC TIER CREATED!")
    print(f"   Slides: {result_basic['slides_created']}")
    print(f"   AI Features: {len(result_basic.get('ai_features_used', []))}")
    print(f"   File Size: {file_size:.1f} KB")
    print(f"   Output: {output_basic}")
else:
    print(f"\n❌ BASIC FAILED: {result_basic.get('error')}")

# ============================================================================
# TIER 2: PRO (AI Summary + Category Insights)
# ============================================================================
print("\n" + "=" * 100)
print("📊 TIER 2: PRO (7 slides, AI summary + category insights)")
print("=" * 100)

converter_pro = ExcelToPPTConverter(
    user_tier='pro',
    user_id='tier_pro',
    user_metadata=user_metadata
)

output_pro = os.path.join(output_dir, "Tech_Stocks_PRO_Tier.pptx")

result_pro = converter_pro.convert_professional(
    excel_path=excel_file,
    output_path=output_pro,
    presentation_title="Tech Stocks Q4 2025 - Professional Analysis",
    user_ppt_count=0
)

if result_pro['success']:
    file_size = os.path.getsize(output_pro) / 1024
    print(f"\n✅ PRO TIER CREATED!")
    print(f"   Slides: {result_pro['slides_created']}")
    print(f"   AI Features: {len(result_pro.get('ai_features_used', []))}")
    print(f"   File Size: {file_size:.1f} KB")
    print(f"   Output: {output_pro}")
else:
    print(f"\n❌ PRO FAILED: {result_pro.get('error')}")

# ============================================================================
# TIER 3: AI_PRO (Full Suite - All AI Features)
# ============================================================================
print("\n" + "=" * 100)
print("📊 TIER 3: AI_PRO (8 slides, full AI suite)")
print("=" * 100)

converter_ai_pro = ExcelToPPTConverter(
    user_tier='ai_pro',
    user_id='tier_ai_pro',
    user_metadata=user_metadata
)

output_ai_pro = os.path.join(output_dir, "Tech_Stocks_AI_PRO_Tier.pptx")

result_ai_pro = converter_ai_pro.convert_professional(
    excel_path=excel_file,
    output_path=output_ai_pro,
    presentation_title="Tech Stocks Q4 2025 - AI-Powered Analysis",
    user_ppt_count=0
)

if result_ai_pro['success']:
    file_size = os.path.getsize(output_ai_pro) / 1024
    print(f"\n✅ AI_PRO TIER CREATED!")
    print(f"   Slides: {result_ai_pro['slides_created']}")
    print(f"   AI Features: {len(result_ai_pro.get('ai_features_used', []))}")
    print(f"   File Size: {file_size:.1f} KB")
    print(f"   Output: {output_ai_pro}")
else:
    print(f"\n❌ AI_PRO FAILED: {result_ai_pro.get('error')}")

# ============================================================================
# SPECIAL: GOOGL FINANCIAL DEEP DIVE
# ============================================================================
print("\n" + "=" * 100)
print("📊 SPECIAL: GOOGL FINANCIAL DEEP DIVE")
print("=" * 100)

googl_metadata = {
    'name': 'Google Financial Analyst',
    'company': 'Alphabet Inc. Investment Research',
    'email': 'research@alphabet-invest.com'
}

converter_googl = ExcelToPPTConverter(
    user_tier='ai_pro',  # Use full features for deep dive
    user_id='googl_analysis',
    user_metadata=googl_metadata
)

output_googl = os.path.join(output_dir, "GOOGL_Financial_Analysis.pptx")

result_googl = converter_googl.convert_professional(
    excel_path=excel_file,
    output_path=output_googl,
    presentation_title="Alphabet (GOOGL) - Comprehensive Financial Analysis Q4 2025",
    user_ppt_count=0
)

if result_googl['success']:
    file_size = os.path.getsize(output_googl) / 1024
    print(f"\n✅ GOOGL ANALYSIS CREATED!")
    print(f"   Slides: {result_googl['slides_created']}")
    print(f"   AI Features: {len(result_googl.get('ai_features_used', []))}")
    print(f"   File Size: {file_size:.1f} KB")
    print(f"   Output: {output_googl}")
    print(f"\n   📈 GOOGL-Specific Data:")
    print(f"      • Stock Price Trends (5-year history)")
    print(f"      • Income Statement Analysis")
    print(f"      • Balance Sheet Health")
    print(f"      • Cash Flow Metrics")
else:
    print(f"\n❌ GOOGL FAILED: {result_googl.get('error')}")

# ============================================================================
# SUMMARY COMPARISON
# ============================================================================
print("\n" + "=" * 100)
print("📊 TIER COMPARISON SUMMARY")
print("=" * 100)

print("\n┌─────────────┬────────┬──────────────┬───────────────────────────────────┐")
print("│ TIER        │ SLIDES │ AI FEATURES  │ KEY FEATURES                      │")
print("├─────────────┼────────┼──────────────┼───────────────────────────────────┤")

if result_basic['success']:
    print(f"│ BASIC       │   {result_basic['slides_created']}    │      {len(result_basic.get('ai_features_used', []))}       │ Executive Summary (AI)            │")
else:
    print("│ BASIC       │   -    │      -       │ FAILED                            │")

if result_pro['success']:
    print(f"│ PRO         │   {result_pro['slides_created']}    │      {len(result_pro.get('ai_features_used', []))}       │ + Category Insights (AI)          │")
else:
    print("│ PRO         │   -    │      -       │ FAILED                            │")

if result_ai_pro['success']:
    print(f"│ AI_PRO      │   {result_ai_pro['slides_created']}    │      {len(result_ai_pro.get('ai_features_used', []))}       │ + Advanced Predictions (AI)       │")
else:
    print("│ AI_PRO      │   -    │      -       │ FAILED                            │")

print("└─────────────┴────────┴──────────────┴───────────────────────────────────┘")

print("\n" + "=" * 100)
print("🎉 GENERATION COMPLETE!")
print("=" * 100)

print("\n📁 Generated Files:")
print(f"   1. {output_basic}")
print(f"   2. {output_pro}")
print(f"   3. {output_ai_pro}")
print(f"   4. {output_googl}")

print("\n💡 What's Different:")
print("   BASIC:   Basic charts + AI executive summary")
print("   PRO:     All charts + AI insights on categories")
print("   AI_PRO:  Everything + AI predictions & anomaly detection")
print("   GOOGL:   Focused on Alphabet financial metrics")

print("\n✨ All presentations include:")
print("   • NO 'nan%' values")
print("   • Clear chart legends")
print("   • Finance theme colors")
print("   • Real Excel data")
