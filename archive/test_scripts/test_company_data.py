"""
Test Professional Slides with Company Bundle Data
Real financial data: AAPL, MSFT, GOOGL, AMZN, TSLA
"""

import sys
import os
from pathlib import Path

# Add project root
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from src.converter.excel_to_ppt_converter import ExcelToPPTConverter

print("=" * 80)
print("🎨 PROFESSIONAL SLIDES WITH REAL COMPANY DATA")
print("=" * 80)

# Setup
excel_file = "examples/Company_Data/company_bundle.xlsx"
output_dir = "examples/professional_demo"
os.makedirs(output_dir, exist_ok=True)

# User metadata
user_metadata = {
    'name': 'Financial Analyst',
    'company': 'Tech Investment Partners',
    'email': 'analyst@techinvest.com'
}

# Test all tiers with real data
tiers_to_test = [
    ('free', 'FREE Tier - Basic Tech Stock Analysis'),
    ('basic', 'BASIC Tier - Tech Stocks with AI Summary'),
    ('pro', 'PRO Tier - Complete Tech Stock Report'),
    ('ai_pro', 'AI PRO Tier - Advanced Tech Stock Analytics')
]

results = []

for tier, description in tiers_to_test:
    print(f"\n{'='*80}")
    print(f"🎯 Testing {tier.upper()} Tier")
    print(f"📊 {description}")
    print(f"{'='*80}")
    
    try:
        # Create converter
        converter = ExcelToPPTConverter(
            user_tier=tier,
            user_id=f'test_{tier}',
            user_metadata=user_metadata
        )
        
        output_file = os.path.join(output_dir, f"tech_stocks_{tier}_tier.pptx")
        
        print(f"\n🔄 Converting with real financial data...")
        
        # Convert
        result = converter.convert_professional(
            excel_path=excel_file,
            output_path=output_file,
            presentation_title="Tech Stocks Performance Analysis - Q4 2025",
            user_ppt_count=0
        )
        
        if result['success']:
            file_size = os.path.getsize(output_file) / 1024
            
            print(f"\n✅ SUCCESS!")
            print(f"   📊 Slides: {result['slides_created']}")
            print(f"   🎨 Template: {result['template_used']}")
            print(f"   🤖 AI Features: {result.get('ai_features_used', [])}")
            print(f"   📦 Size: {file_size:.1f} KB")
            print(f"   💾 File: {output_file}")
            
            if result.get('errors'):
                print(f"\n   ⚠️  Errors:")
                for error in result['errors'][:3]:
                    print(f"      - {error}")
            
            results.append({
                'tier': tier.upper(),
                'success': True,
                'slides': result['slides_created'],
                'ai_features': len(result.get('ai_features_used', [])),
                'file_size': file_size,
                'file': output_file
            })
        else:
            print(f"\n❌ FAILED: {result.get('error')}")
            results.append({
                'tier': tier.upper(),
                'success': False,
                'error': result.get('error')
            })
    
    except Exception as e:
        print(f"\n❌ ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        results.append({
            'tier': tier.upper(),
            'success': False,
            'error': str(e)
        })

# Print summary
print(f"\n\n{'='*80}")
print("📊 TECH STOCKS ANALYSIS - RESULTS SUMMARY")
print(f"{'='*80}\n")

print(f"{'Tier':<12} {'Status':<10} {'Slides':<8} {'AI Features':<12} {'Size (KB)':<12}")
print(f"{'-'*80}")

for result in results:
    tier = result['tier']
    if result['success']:
        status = "✅ Pass"
        slides = str(result['slides'])
        ai = str(result['ai_features'])
        size = f"{result['file_size']:.1f}"
    else:
        status = "❌ Fail"
        slides = "-"
        ai = "-"
        size = "-"
    
    print(f"{tier:<12} {status:<10} {slides:<8} {ai:<12} {size:<12}")

print(f"\n{'='*80}")
print("📁 Generated Files:")
print(f"{'='*80}")
for result in results:
    if result['success']:
        print(f"   📄 {result['file']}")

print(f"\n{'='*80}")
print("✨ Tech Stocks Professional Presentations Complete!")
print(f"{'='*80}")

success_count = sum(1 for r in results if r['success'])
print(f"\n✅ Successful: {success_count}/{len(results)}")
print(f"❌ Failed: {len(results) - success_count}/{len(results)}")

print("\n💡 These presentations now include:")
print("   • Real stock price charts (AAPL, MSFT, GOOGL, AMZN, TSLA)")
print("   • Financial statement data (Income, Balance Sheet, Cash Flow)")
print("   • AI-powered insights and analysis")
print("   • Professional KPI cards with real metrics")
print("   • Category comparison charts")
print(f"   • Branded for {user_metadata['company']}")
