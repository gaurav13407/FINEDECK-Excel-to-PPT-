"""
Comprehensive Test Script: Generate PPTs for All Excel Files Using All 4 Tiers
This demonstrates the differences between Free, Basic, Pro, and AI Pro tiers.
"""

import os
import sys
from pathlib import Path

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent))

from src.converter.excel_to_ppt_converter import convert_excel_to_ppt, TIER_CONFIG

def test_all_tiers_all_files():
    """
    Test all Excel files with all 4 subscription tiers.
    Creates organized output folders showing tier differences.
    """
    
    # Define test files
    test_files = [
        {
            "name": "Company Bundle",
            "path": "examples/Company_Data/company_bundle.xlsx",
            "description": "Company financial data"
        },
        {
            "name": "Portfolio Allocation",
            "path": "examples/Portfolio Allocation Data.xlsx",
            "description": "Investment portfolio data"
        },
        {
            "name": "Risk Metrics",
            "path": "examples/Risk Metrics Data.xlsx",
            "description": "Financial risk analysis"
        },
        {
            "name": "PnL Sample",
            "path": "examples/Sample_pnl.xlsx",
            "description": "Profit & Loss statement"
        },
        {
            "name": "Scenario Comparison",
            "path": "examples/Scenario Comparison (Bull_Bear_Base).xlsx",
            "description": "Bull/Bear/Base scenario analysis"
        },
        {
            "name": "Finance Sample",
            "path": "examples/finance_sample.xlsx",
            "description": "General finance data"
        }
    ]
    
    # Define all 4 tiers with their characteristics
    tiers = [
        {
            "name": "free",
            "display": "FREE",
            "features": "1 sheet, Basic formatting, No AI",
            "color": "🆓"
        },
        {
            "name": "basic",
            "display": "BASIC",
            "features": "5 sheets, AI Titles, Basic templates",
            "color": "🥉"
        },
        {
            "name": "pro",
            "display": "PRO",
            "features": "20 sheets, AI Titles + Templates, 12 templates",
            "color": "🥈"
        },
        {
            "name": "ai_pro",
            "display": "AI PRO",
            "features": "Unlimited sheets, All 6 AI features, 12 templates",
            "color": "🥇"
        }
    ]
    
    # Create output directory
    output_base = Path("examples/tier_comparison_demo")
    output_base.mkdir(exist_ok=True)
    
    print("=" * 80)
    print("🎯 FINEDECK TIER COMPARISON TEST")
    print("=" * 80)
    print(f"\n📊 Testing {len(test_files)} Excel files with {len(tiers)} subscription tiers")
    print(f"📁 Output directory: {output_base}\n")
    
    # Statistics
    total_tests = len(test_files) * len(tiers)
    completed = 0
    failed = 0
    results = []
    
    # Test each file with each tier
    for file_info in test_files:
        file_path = Path(file_info["path"])
        
        # Skip if file doesn't exist
        if not file_path.exists():
            print(f"⚠️  Skipping {file_info['name']}: File not found")
            failed += len(tiers)
            continue
        
        print(f"\n{'=' * 80}")
        print(f"📄 Testing File: {file_info['name']}")
        print(f"📝 Description: {file_info['description']}")
        print(f"📂 Path: {file_path}")
        print(f"{'=' * 80}\n")
        
        # Create folder for this file
        file_output_dir = output_base / file_path.stem
        file_output_dir.mkdir(exist_ok=True)
        
        # Test with each tier
        for tier_info in tiers:
            tier_name = tier_info["name"]
            tier_display = tier_info["display"]
            tier_icon = tier_info["color"]
            
            print(f"\n{tier_icon} Testing with {tier_display} tier...")
            print(f"   Features: {tier_info['features']}")
            
            try:
                # Output filename
                output_file = file_output_dir / f"{file_path.stem}_{tier_name.upper()}.pptx"
                
                # Convert using the tier-based converter
                print(f"   🔄 Converting...")
                result = convert_excel_to_ppt(
                    excel_path=str(file_path),
                    output_path=str(output_file),
                    user_tier=tier_name,  # Set the subscription tier
                    user_ppt_count=0  # Demo user with no prior usage
                )
                
                # Display results
                if result["success"]:
                    print(f"   ✅ SUCCESS!")
                    print(f"   📊 Slides created: {result.get('slides_created', 0)}")
                    print(f"   🎨 Template used: {result.get('template_used', 'N/A')}")
                    print(f"   🤖 AI features: {', '.join(result.get('ai_features_used', [])) or 'None'}")
                    print(f"   💾 Output: {output_file.name}")
                    
                    # Get file size
                    file_size = output_file.stat().st_size / 1024  # KB
                    print(f"   📦 File size: {file_size:.1f} KB")
                    
                    completed += 1
                    
                    results.append({
                        "file": file_info['name'],
                        "tier": tier_display,
                        "status": "✅ Success",
                        "sheets": result.get('slides_created', 0) - 1,  # -1 for title slide
                        "charts": result.get('slides_created', 0) - 1,
                        "ai_features": len(result.get('ai_features_used', [])),
                        "size_kb": f"{file_size:.1f}"
                    })
                else:
                    print(f"   ❌ FAILED: {result.get('error', 'Unknown error')}")
                    failed += 1
                    
                    results.append({
                        "file": file_info['name'],
                        "tier": tier_display,
                        "status": "❌ Failed",
                        "sheets": 0,
                        "charts": 0,
                        "ai_features": 0,
                        "size_kb": "0"
                    })
                    
            except Exception as e:
                print(f"   ❌ ERROR: {str(e)}")
                failed += 1
                
                results.append({
                    "file": file_info['name'],
                    "tier": tier_display,
                    "status": f"❌ Error",
                    "sheets": 0,
                    "charts": 0,
                    "ai_features": 0,
                    "size_kb": "0"
                })
    
    # Print summary
    print("\n" + "=" * 80)
    print("📊 TEST SUMMARY")
    print("=" * 80)
    
    print(f"\n✅ Completed: {completed}/{total_tests}")
    print(f"❌ Failed: {failed}/{total_tests}")
    print(f"📈 Success Rate: {(completed/total_tests)*100:.1f}%\n")
    
    # Print detailed results table
    print("=" * 80)
    print("📋 DETAILED RESULTS")
    print("=" * 80)
    print(f"\n{'File':<25} {'Tier':<10} {'Status':<12} {'Sheets':<8} {'Charts':<8} {'AI':<6} {'Size':<10}")
    print("-" * 90)
    
    for result in results:
        print(f"{result['file']:<25} {result['tier']:<10} {result['status']:<12} "
              f"{result['sheets']:<8} {result['charts']:<8} {result['ai_features']:<6} {result['size_kb']:<10} KB")
    
    # Tier comparison
    print("\n" + "=" * 80)
    print("🎯 TIER FEATURE COMPARISON")
    print("=" * 80)
    
    print("\n🆓 FREE Tier:")
    print("   - Process: 1 sheet only")
    print("   - AI Features: None")
    print("   - Templates: 1 basic template")
    print("   - Best for: Quick single-sheet presentations")
    
    print("\n🥉 BASIC Tier ($25/month):")
    print("   - Process: Up to 5 sheets")
    print("   - AI Features: AI-generated titles")
    print("   - Templates: 1 basic template")
    print("   - Monthly Limit: 7 PPTs")
    print("   - Best for: Small businesses, basic reporting")
    
    print("\n🥈 PRO Tier ($49/month):")
    print("   - Process: Up to 20 sheets")
    print("   - AI Features: AI titles + Smart templates")
    print("   - Templates: 12 professional templates")
    print("   - Monthly Limit: 15 PPTs")
    print("   - Best for: Professional analysts, consultants")
    
    print("\n🥇 AI PRO Tier ($99/month):")
    print("   - Process: Unlimited sheets")
    print("   - AI Features: All 6 AI features")
    print("   - Templates: 12 professional templates")
    print("   - Monthly Limit: Unlimited PPTs")
    print("   - Best for: Enterprises, heavy users, data teams")
    
    print("\n" + "=" * 80)
    print("📁 Output Location:")
    print(f"   {output_base.absolute()}")
    print("=" * 80)
    
    print("\n✨ Test Complete! Check the output folders to compare tier differences.\n")
    
    return completed, failed, total_tests

if __name__ == "__main__":
    print("\n🚀 Starting Comprehensive Tier Testing...\n")
    
    try:
        completed, failed, total = test_all_tiers_all_files()
        
        if completed == total:
            print("🎉 All tests passed successfully!")
            sys.exit(0)
        elif completed > 0:
            print(f"⚠️  Some tests failed ({failed}/{total})")
            sys.exit(1)
        else:
            print("❌ All tests failed!")
            sys.exit(1)
            
    except KeyboardInterrupt:
        print("\n\n⚠️  Test interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ Fatal error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
