"""
Quick Demo: Generate PPTs for Key Excel Files Using All 4 Tiers
Shows differences between Free, Basic, Pro, and AI Pro tiers
"""

import os
import sys
from pathlib import Path

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent))

from src.converter.excel_to_ppt_converter import convert_excel_to_ppt, TIER_CONFIG

def quick_demo():
    """Quick demo with 2 key files and all 4 tiers"""
    
    # Test just 2 representative files
    test_files = [
        {
            "name": "Company Bundle",
            "path": "examples/Company_Data/company_bundle.xlsx",
        },
        {
            "name": "PnL Sample",
            "path": "examples/Sample_pnl.xlsx",
        }
    ]
    
    # All 4 tiers
    tiers = ["free", "basic", "pro", "ai_pro"]
    tier_display = {"free": "🆓 FREE", "basic": "🥉 BASIC", "pro": "🥈 PRO", "ai_pro": "🥇 AI PRO"}
    
    # Create output directory
    output_base = Path("examples/quick_demo")
    output_base.mkdir(exist_ok=True)
    
    print("=" * 80)
    print("🎯 FINEDECK QUICK DEMO - ALL TIERS")
    print("=" * 80)
    print(f"\n📊 Testing 2 files × 4 tiers = 8 presentations")
    print(f"📁 Output: {output_base}\n")
    
    results = []
    
    for file_info in test_files:
        file_path = Path(file_info["path"])
        
        if not file_path.exists():
            print(f"⚠️  Skipping {file_info['name']}: File not found\n")
            continue
        
        print(f"\n{'─' * 80}")
        print(f"📄 {file_info['name']}")
        print(f"{'─' * 80}")
        
        # Create folder for this file
        file_output_dir = output_base / file_path.stem
        file_output_dir.mkdir(exist_ok=True)
        
        # Test with each tier
        for tier in tiers:
            tier_name = tier_display[tier]
            config = TIER_CONFIG[tier]
            
            print(f"\n{tier_name}")
            print(f"  • Max sheets: {config['max_sheets']}")
            print(f"  • AI features: {', '.join(config['ai_features']) or 'None'}")
            print(f"  • Templates: {len(config['templates'])}")
            
            try:
                # Output filename
                output_file = file_output_dir / f"{file_path.stem}_{tier.upper()}.pptx"
                
                # Convert
                print(f"  🔄 Converting...", end=" ", flush=True)
                result = convert_excel_to_ppt(
                    excel_path=str(file_path),
                    output_path=str(output_file),
                    user_tier=tier,
                    user_ppt_count=0
                )
                
                if result["success"]:
                    file_size = output_file.stat().st_size / 1024
                    slides = result.get('slides_created', 0)
                    print(f"✅ {slides} slides, {file_size:.1f} KB")
                    
                    results.append({
                        "file": file_info['name'],
                        "tier": tier_name,
                        "slides": slides,
                        "size": f"{file_size:.1f} KB",
                        "status": "✅"
                    })
                else:
                    print(f"❌ {result.get('error', 'Failed')}")
                    results.append({
                        "file": file_info['name'],
                        "tier": tier_name,
                        "slides": 0,
                        "size": "0 KB",
                        "status": "❌"
                    })
                    
            except Exception as e:
                print(f"❌ Error: {str(e)[:50]}")
                results.append({
                    "file": file_info['name'],
                    "tier": tier_name,
                    "slides": 0,
                    "size": "0 KB",
                    "status": "❌"
                })
    
    # Summary
    print("\n" + "=" * 80)
    print("📊 RESULTS SUMMARY")
    print("=" * 80)
    print(f"\n{'File':<20} {'Tier':<12} {'Slides':<8} {'Size':<12} {'Status':<8}")
    print("─" * 80)
    
    for r in results:
        print(f"{r['file']:<20} {r['tier']:<12} {r['slides']:<8} {r['size']:<12} {r['status']:<8}")
    
    # Feature comparison
    print("\n" + "=" * 80)
    print("🎯 TIER COMPARISON")
    print("=" * 80)
    
    print("\n🆓 FREE:")
    print("   • 1 sheet only")
    print("   • No AI features")
    print("   • Basic template")
    print("   • 1 PPT/month limit")
    
    print("\n🥉 BASIC ($25/month):")
    print("   • Up to 5 sheets")
    print("   • AI-generated titles")
    print("   • Basic template")
    print("   • 7 PPTs/month")
    
    print("\n🥈 PRO ($49/month):")
    print("   • Up to 20 sheets")
    print("   • AI titles + Smart templates")
    print("   • 12 professional templates")
    print("   • 15 PPTs/month")
    
    print("\n🥇 AI PRO ($99/month):")
    print("   • Unlimited sheets")
    print("   • All 6 AI features")
    print("   • 12 professional templates")
    print("   • Unlimited PPTs/month")
    
    print("\n" + "=" * 80)
    print(f"📁 Check output: {output_base.absolute()}")
    print("=" * 80)
    print("\n✅ Demo complete!\n")

if __name__ == "__main__":
    try:
        quick_demo()
    except KeyboardInterrupt:
        print("\n⚠️  Interrupted by user")
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()
