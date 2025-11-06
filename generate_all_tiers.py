"""
Generate PPTs for ALL TIERS with ENHANCED Professional Builder
- BASIC tier (8+ slides, AI summary + smart charts)
- PRO tier (8+ slides, AI summary + insights + smart charts)
- AI_PRO tier (9 slides, full AI features + AI-powered charts)
- Uses EnhancedProfessionalBuilder with SmartChartAnalyzer integration
"""

import sys
import os
from pathlib import Path

project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from src.converter.excel_to_ppt_converter import ExcelToPPTConverter

print("=" * 100)
print("🎯 GENERATING ALL TIERS WITH ENHANCED PROFESSIONAL BUILDER")
print("   ✨ 8+ comprehensive slides per tier")
print("   ✨ AI-powered chart recommendations")
print("   ✨ SmartChartAnalyzer integration")
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
# TIER 1: BASIC (Enhanced 8+ slides with smart charts)
# ============================================================================
print("\n" + "=" * 100)
print("📊 TIER 1: BASIC (8+ slides)")
print("   📄 Executive Summary, Key Metrics, Data Insights, Sector Distribution,")
print("   📄 Key Data Insights, Top Performers, Trend Analysis, Closing")
print("   🤖 AI executive summary + Smart chart selection")
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
    user_ppt_count=0,
    use_professional_structure=True  # Force enhanced 8+ slide structure
)

if result_basic['success']:
    file_size = os.path.getsize(output_basic) / 1024
    print(f"\n✅ BASIC TIER CREATED!")
    print(f"   📊 Slides: {result_basic['slides_created']}")
    print(f"   📄 Slide Names: {', '.join(result_basic.get('slide_names', []))}")
    print(f"   🤖 AI Features: {len(result_basic.get('ai_features_used', []))}")
    print(f"   📂 File Size: {file_size:.1f} KB")
    print(f"   💾 Output: {output_basic}")
else:
    print(f"\n❌ BASIC FAILED: {result_basic.get('error')}")

# ============================================================================
# TIER 2: PRO (Enhanced 8+ slides with AI insights)
# ============================================================================
print("\n" + "=" * 100)
print("📊 TIER 2: PRO (8+ slides)")
print("   📄 Executive Summary, Key Metrics, Data Insights, Sector Distribution,")
print("   📄 Key Data Insights, Top Performers, Trend Analysis, Closing")
print("   🤖 AI summary + insights + Smart charts + Multi-series analysis")
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
    user_ppt_count=0,
    use_professional_structure=True  # Force enhanced 8+ slide structure
)

if result_pro['success']:
    file_size = os.path.getsize(output_pro) / 1024
    print(f"\n✅ PRO TIER CREATED!")
    print(f"   📊 Slides: {result_pro['slides_created']}")
    print(f"   📄 Slide Names: {', '.join(result_pro.get('slide_names', []))}")
    print(f"   🤖 AI Features: {len(result_pro.get('ai_features_used', []))}")
    print(f"   📂 File Size: {file_size:.1f} KB")
    print(f"   💾 Output: {output_pro}")
else:
    print(f"\n❌ PRO FAILED: {result_pro.get('error')}")

# ============================================================================
# TIER 3: AI_PRO (Full Suite - All AI Features + AI-Powered Charts)
# ============================================================================
print("\n" + "=" * 100)
print("📊 TIER 3: AI_PRO (9 comprehensive slides)")
print("   📄 Executive Summary, Key Metrics, Data Insights, Sector Distribution,")
print("   📄 Key Data Insights, Top Performers, Trend Analysis, Closing")
print("   🤖 FULL AI SUITE: AI chart recommendations + AI insights + SmartChartAnalyzer")
print("   📊 Advanced: Multi-series trends, AI-guided chart types, Deep data analysis")
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
    user_ppt_count=0,
    use_professional_structure=True  # Force enhanced 8+ slide structure
)

if result_ai_pro['success']:
    file_size = os.path.getsize(output_ai_pro) / 1024
    print(f"\n✅ AI_PRO TIER CREATED!")
    print(f"   📊 Slides: {result_ai_pro['slides_created']}")
    print(f"   📄 Slide Names: {', '.join(result_ai_pro.get('slide_names', []))}")
    print(f"   🤖 AI Features Used: {', '.join(result_ai_pro.get('ai_features_used', []))}")
    print(f"   📈 Chart Types: AI-recommended bar, line, pie charts")
    print(f"   📂 File Size: {file_size:.1f} KB")
    print(f"   💾 Output: {output_ai_pro}")
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
    user_ppt_count=0,
    use_professional_structure=True  # Force enhanced 8+ slide structure with AI charts
)

if result_googl['success']:
    file_size = os.path.getsize(output_googl) / 1024
    print(f"\n✅ GOOGL ANALYSIS CREATED!")
    print(f"   📊 Slides: {result_googl['slides_created']}")
    print(f"   📄 Slide Names: {', '.join(result_googl.get('slide_names', []))}")
    print(f"   🤖 AI Features: {', '.join(result_googl.get('ai_features_used', []))}")
    print(f"   📈 Advanced Charts: AI-recommended + SmartChartAnalyzer")
    print(f"   📂 File Size: {file_size:.1f} KB")
    print(f"   💾 Output: {output_googl}")
    print(f"\n   � Enhanced Data Visualization:")
    print(f"      • Multi-series trend charts (up to 3 metrics)")
    print(f"      • AI-powered chart type selection")
    print(f"      • Top 10 performers with rankings")
    print(f"      • Sector distribution with 8 categories")
    print(f"      • Deep dive insights with detailed analysis")
else:
    print(f"\n❌ GOOGL FAILED: {result_googl.get('error')}")

# ============================================================================
# SUMMARY COMPARISON
# ============================================================================
print("\n" + "=" * 100)
print("📊 TIER COMPARISON SUMMARY - Enhanced Professional Builder")
print("=" * 100)

print("\n┌─────────────┬────────┬──────────────┬───────────────────────────────────┐")
print("│ TIER        │ SLIDES │ AI FEATURES  │ KEY FEATURES                      │")
print("├─────────────┼────────┼──────────────┼───────────────────────────────────┤")

if result_basic['success']:
    print(f"│ BASIC       │   {result_basic['slides_created']}    │      {len(result_basic.get('ai_features_used', []))}       │ 8+ Slides + AI Summary            │")
else:
    print("│ BASIC       │   -    │      -       │ FAILED                            │")

if result_pro['success']:
    print(f"│ PRO         │   {result_pro['slides_created']}    │      {len(result_pro.get('ai_features_used', []))}       │ + AI Charts + Smart Analysis      │")
else:
    print("│ PRO         │   -    │      -       │ FAILED                            │")

if result_ai_pro['success']:
    print(f"│ AI_PRO      │   {result_ai_pro['slides_created']}    │      {len(result_ai_pro.get('ai_features_used', []))}       │ + Full AI Suite + Deep Dive       │")
else:
    print("│ AI_PRO      │   -    │      -       │ FAILED                            │")

if result_googl['success']:
    print(f"│ GOOGL       │   {result_googl['slides_created']}    │      {len(result_googl.get('ai_features_used', []))}       │ + Financial Deep Dive             │")
else:
    print("│ GOOGL       │   -    │      -       │ FAILED                            │")

print("└─────────────┴────────┴──────────────┴───────────────────────────────────┘")

print("\n" + "=" * 100)
print("🎉 GENERATION COMPLETE!")
print("=" * 100)

print("\n📁 Generated Files:")
print(f"   1. {output_basic}")
print(f"   2. {output_pro}")
print(f"   3. {output_ai_pro}")
print(f"   4. {output_googl}")

print("\n💡 Enhanced Professional Builder Features:")
print("   BASIC:   8-9 comprehensive slides + AI executive summary")
print("   PRO:     8-9 slides + AI charts + SmartChartAnalyzer")
print("   AI_PRO:  9 slides + Full AI suite + AI-recommended charts")
print("   GOOGL:   9 slides + Financial deep dive + Multi-series trends")

print("\n✨ All presentations include:")
print("   • 8+ Professional slides (Title, Summary, Metrics, Insights, etc.)")
print("   • AI-powered chart recommendations")
print("   • SmartChartAnalyzer intelligent chart selection")
print("   • Multi-series trend charts (up to 3 metrics)")
print("   • Top 10 performers with rankings")
print("   • Sector distribution (8 categories)")
print("   • Deep dive insights with detailed analysis")
print("   • Finance theme colors & professional formatting")
print("   • Real Excel data with NO 'nan%' values")
