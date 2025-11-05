"""
Quick test to verify professional slide builder imports and basic functionality
"""

import sys
from pathlib import Path

# Add project root
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

print("Testing imports...")

try:
    from src.converter.professional_slide_builder import ProfessionalSlideBuilder, SLIDE_STRUCTURE, TIER_SLIDE_CONFIG
    print("✅ ProfessionalSlideBuilder imported successfully")
    
    print(f"\n📊 Slide Structure Configuration:")
    for slide_key, slide_info in SLIDE_STRUCTURE.items():
        print(f"   - {slide_info['name']}: {'Required' if slide_info['required'] else 'Optional'}")
    
    print(f"\n🎯 Tier Configurations:")
    for tier, config in TIER_SLIDE_CONFIG.items():
        print(f"   - {tier.upper()}: {config['min_slides']}-{config['max_slides']} slides")
    
except Exception as e:
    print(f"❌ Import error: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

try:
    from src.converter.excel_to_ppt_converter import ExcelToPPTConverter
    print("\n✅ ExcelToPPTConverter imported successfully")
    
    # Test converter creation
    converter = ExcelToPPTConverter(user_tier='free')
    print(f"✅ Converter created for FREE tier")
    
    # Test with metadata
    metadata = {'name': 'Test User', 'company': 'Test Company'}
    converter = ExcelToPPTConverter(user_tier='ai_pro', user_metadata=metadata)
    print(f"✅ Converter created for AI_PRO tier with metadata")
    
except Exception as e:
    print(f"❌ Converter error: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

print("\n✨ All imports successful! Ready to test professional slides.")
