"""
Test Professional 7-Slide Structure for All Tiers
Generate sample presentations to verify:
- FREE: 5-7 slides, no AI features
- BASIC: 7 slides with basic AI (executive summary)
- PRO: 8-9 slides with enhanced AI (executive summary + category insights)
- AI_PRO: 9-10 slides with full AI (all AI features)
"""

import os
import sys
from pathlib import Path

# Add project root to path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from src.converter.excel_to_ppt_converter import ExcelToPPTConverter


def test_professional_structure():
    """Test professional slide structure for all tiers"""
    
    # Test data
    excel_file = "examples/Sample_pnl.xlsx"
    output_dir = "examples/professional_demo"
    
    # Create output directory
    os.makedirs(output_dir, exist_ok=True)
    
    # User metadata for branding
    user_metadata = {
        'name': 'John Doe',
        'company': 'FinDeck Analytics Inc.',
        'email': 'john.doe@findeck.com'
    }
    
    # Test all tiers
    tiers = ['free', 'basic', 'pro', 'ai_pro']
    
    print("=" * 80)
    print("🎨 PROFESSIONAL SLIDE STRUCTURE TEST")
    print("=" * 80)
    print(f"\n📁 Input File: {excel_file}")
    print(f"📂 Output Directory: {output_dir}")
    print(f"👤 User: {user_metadata['name']} ({user_metadata['company']})")
    print("\n")
    
    results = []
    
    for tier in tiers:
        print(f"\n{'='*80}")
        print(f"🎯 Testing {tier.upper()} Tier")
        print(f"{'='*80}")
        
        try:
            # Create converter for this tier
            converter = ExcelToPPTConverter(
                user_tier=tier,
                user_id='test_user_123',
                user_metadata=user_metadata
            )
            
            # Generate output filename
            output_file = os.path.join(output_dir, f"professional_{tier}_tier.pptx")
            
            # Convert with professional structure
            result = converter.convert_professional(
                excel_path=excel_file,
                output_path=output_file,
                presentation_title="Q4 Financial Performance Report",
                user_ppt_count=0,
                use_professional_structure=True
            )
            
            if result['success']:
                file_size = os.path.getsize(output_file) / 1024  # KB
                
                tier_result = {
                    'tier': tier.upper(),
                    'success': True,
                    'slides': result['slides_created'],
                    'template': result['template_used'],
                    'ai_features': len(result.get('ai_features_used', [])),
                    'file_size_kb': file_size,
                    'output': output_file,
                    'errors': len(result.get('errors', []))
                }
                
                print(f"\n✅ SUCCESS")
                print(f"   📊 Slides Created: {result['slides_created']}")
                print(f"   🎨 Template: {result['template_used']}")
                print(f"   🤖 AI Features: {result.get('ai_features_used', [])}")
                print(f"   📦 File Size: {file_size:.1f} KB")
                print(f"   💾 Output: {output_file}")
                
                if result.get('errors'):
                    print(f"   ⚠️  Errors: {len(result['errors'])}")
                    for error in result['errors'][:3]:  # Show first 3 errors
                        print(f"      - {error}")
                
            else:
                tier_result = {
                    'tier': tier.upper(),
                    'success': False,
                    'error': result.get('error', 'Unknown error')
                }
                print(f"\n❌ FAILED: {result.get('error')}")
            
            results.append(tier_result)
            
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
    print("📊 SUMMARY - PROFESSIONAL SLIDE STRUCTURE")
    print(f"{'='*80}\n")
    
    print(f"{'Tier':<12} {'Status':<10} {'Slides':<8} {'AI Features':<12} {'Size (KB)':<12}")
    print(f"{'-'*80}")
    
    for result in results:
        tier = result['tier']
        if result['success']:
            status = "✅ Pass"
            slides = str(result['slides'])
            ai_features = str(result['ai_features'])
            size = f"{result['file_size_kb']:.1f}"
        else:
            status = "❌ Fail"
            slides = "-"
            ai_features = "-"
            size = "-"
        
        print(f"{tier:<12} {status:<10} {slides:<8} {ai_features:<12} {size:<12}")
    
    # Verify slide count expectations
    print(f"\n{'='*80}")
    print("🎯 SLIDE COUNT VERIFICATION")
    print(f"{'='*80}\n")
    
    expectations = {
        'FREE': (5, 7, "No AI, basic slides only"),
        'BASIC': (7, 7, "AI executive summary"),
        'PRO': (8, 9, "AI summary + category insights"),
        'AI_PRO': (9, 10, "Full AI: summary + insights + predictions")
    }
    
    for result in results:
        if result['success']:
            tier = result['tier']
            slides = result['slides']
            min_slides, max_slides, description = expectations[tier]
            
            if min_slides <= slides <= max_slides:
                status = "✅ Correct"
            else:
                status = f"⚠️  Expected {min_slides}-{max_slides}"
            
            print(f"{tier:<12} {slides:<8} slides - {status}")
            print(f"             {description}")
            print()
    
    print(f"\n{'='*80}")
    print("✨ Professional Slide Structure Test Complete!")
    print(f"{'='*80}\n")
    
    # Success count
    success_count = sum(1 for r in results if r['success'])
    print(f"✅ Successful: {success_count}/{len(results)}")
    print(f"❌ Failed: {len(results) - success_count}/{len(results)}")
    
    return results


if __name__ == "__main__":
    test_professional_structure()
