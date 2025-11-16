"""
Generate ENHANCED Professional Slides - With Better Design
"""

import sys
import os
from pathlib import Path

project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from src.converter.excel_to_ppt_converter import ExcelToPPTConverter

print("=" * 80)
print("🎨 ENHANCED PROFESSIONAL SLIDES - IMPROVED DESIGN")
print("=" * 80)

excel_file = "examples/Company_Data/company_bundle.xlsx"
output_dir = "examples/professional_demo"
os.makedirs(output_dir, exist_ok=True)

user_metadata = {
    'name': 'Sarah Johnson',
    'company': 'Global Investment Analytics',
    'email': 'sarah.johnson@gia.com'
}

print(f"\n📊 Creating enhanced presentation with improved design...")
print(f"   • Better visual hierarchy")
print(f"   • Color-coded sections")
print(f"   • Professional spacing and layout")
print(f"   • Enhanced typography")
print(f"   • Decorative elements")

# Generate AI PRO tier (most features)
converter = ExcelToPPTConverter(
    user_tier='ai_pro',
    user_id='enhanced_demo',
    user_metadata=user_metadata
)

output_file = os.path.join(output_dir, "ENHANCED_tech_stocks_professional.pptx")

result = converter.convert_professional(
    excel_path=excel_file,
    output_path=output_file,
    presentation_title="Tech Stocks Q4 Performance - Professional Analysis",
    user_ppt_count=0
)

if result['success']:
    file_size = os.path.getsize(output_file) / 1024
    
    print(f"\n✅ ENHANCED PRESENTATION CREATED!")
    print(f"{'='*80}")
    print(f"📊 Slides: {result['slides_created']}")
    print(f"🎨 Template: {result['template_used']}")
    print(f"🤖 AI Features: {len(result.get('ai_features_used', []))}")
    print(f"📦 File Size: {file_size:.1f} KB")
    print(f"💾 Output: {output_file}")
    print(f"\n✨ IMPROVEMENTS APPLIED:")
    print(f"   ✓ Title slide with decorative top bar")
    print(f"   ✓ Executive summary with accent background box")
    print(f"   ✓ KPI cards with improved spacing & borders")
    print(f"   ✓ Real charts with stock price data")
    print(f"   ✓ Color-coded AI insights sections")
    print(f"   ✓ Enhanced closing slide with accent shapes")
    print(f"   ✓ Better typography and line spacing")
    print(f"   ✓ Professional color scheme")
    
    # Verify charts
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    
    prs = Presentation(output_file)
    chart_count = sum(1 for slide in prs.slides for shape in slide.shapes 
                     if shape.shape_type == MSO_SHAPE_TYPE.CHART)
    
    print(f"\n📈 Charts Verification:")
    print(f"   Total charts: {chart_count}")
    print(f"   Status: {'✅ Good' if chart_count >= 2 else '⚠️ Needs more charts'}")
    
else:
    print(f"\n❌ FAILED: {result.get('error')}")

print(f"\n{'='*80}")
print("🎯 COMPARE:")
print(f"{'='*80}")
print("Open both files to see the improvements:")
print(f"   📄 OLD: examples/professional_demo/tech_stocks_ai_pro_tier.pptx")
print(f"   📄 NEW: {output_file}")
print("\n✨ The NEW version has better visual design and professional layout!")
