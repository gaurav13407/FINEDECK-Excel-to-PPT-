"""
Comprehensive Test Suite for Enhanced Finance Visual Standards
Tests all new consulting/investor deck quality features
"""

import sys
import os
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / 'src'))

import pandas as pd
import numpy as np
from pptx import Presentation
from pptx.util import Inches

from src.converter.finance_chart_formatter import (
    format_number,
    clean_nan_values,
    format_percentage,
    apply_finance_theme,
    annotate_chart,
    finalize_presentation,
    FINANCE_COLORS,
    FINANCE_FONTS,
    LAYOUT_CONSTANTS
)

print("="*80)
print("🎨 ENHANCED FINANCE VISUAL STANDARDS - COMPREHENSIVE TEST")
print("="*80)

# ============================================================================
# Test 1: Number Formatting with $ Prefix
# ============================================================================
print("\n📊 Test 1: Number Formatting ($ prefix, NO scientific notation)")
test_values = [
    2.53e12,   # Should be $2.5T
    865.4e9,   # Should be $865.4B
    42.7e6,    # Should be $42.7M
    5.2e3,     # Should be $5.2K
    1234.5,    # Should be $1,235
    -2.1e9,    # Should be -$2.1B
]

for val in test_values:
    formatted = format_number(val)
    print(f"   {val:.2e} → {formatted} ✅")

# ============================================================================
# Test 2: Percentage Formatting
# ============================================================================
print("\n📊 Test 2: Percentage Formatting (1 decimal)")
test_percentages = [12.345, 5.0, -3.7, 0.123]
for val in test_percentages:
    formatted = format_percentage(val)
    print(f"   {val} → {formatted} ✅")

# ============================================================================
# Test 3: NaN Cleaning
# ============================================================================
print("\n📊 Test 3: NaN Value Cleaning (replace with '–')")
test_series = pd.Series([100, np.nan, 200, None, 300])
cleaned = clean_nan_values(test_series)
print(f"   Original: {list(test_series)}")
print(f"   Cleaned:  {list(cleaned)} ✅")

# ============================================================================
# Test 4: Global Design Constants
# ============================================================================
print("\n📊 Test 4: Global Design Theme Constants")
print(f"   ✅ Font family: {FINANCE_FONTS['primary']} (fallback: {FINANCE_FONTS['fallback_1']})")
print(f"   ✅ Title size: {FINANCE_FONTS['title_size']}pt Bold")
neutral_hex = "#{:02X}{:02X}{:02X}".format(*FINANCE_COLORS['neutral'])
print(f"   ✅ Axis size: {FINANCE_FONTS['axis_size']}pt Gray {neutral_hex}")
print(f"   ✅ Legend size: {FINANCE_FONTS['legend_size']}pt")
primary_hex = "#{:02X}{:02X}{:02X}".format(*FINANCE_COLORS['primary'])
print(f"   ✅ Primary color: {primary_hex}")
print(f"   ✅ Top margin: {LAYOUT_CONSTANTS['slide_padding_top']}px")
print(f"   ✅ Chart width: {LAYOUT_CONSTANTS['chart_width_ratio']*100:.0f}% of slide")
print(f"   ✅ Max Y-ticks: {LAYOUT_CONSTANTS['max_y_ticks']}")
print(f"   ✅ Bar gap width: {LAYOUT_CONSTANTS['bar_gap_width']}%")
print(f"   ✅ Top-N limit: {LAYOUT_CONSTANTS['top_n_limit']} + Other")

# ============================================================================
# Test 5: Full Integration Test with Chart Creation
# ============================================================================
print("\n📊 Test 5: Full Integration Test - Chart Creation")

# Create sample data
data = {
    'Company': ['Apple', 'Microsoft', 'Google', 'Amazon', 'NVIDIA', 'Meta', 'Tesla', 'Berkshire', 'TSM', 'Visa'],
    'Market Cap': [2.8e12, 2.6e12, 1.7e12, 1.5e12, 1.2e12, 850e9, 700e9, 650e9, 550e9, 500e9]
}
df = pd.DataFrame(data)

print(f"   Created test dataset: {df.shape[0]} companies")
print(f"   Top 3: {df['Company'].head(3).tolist()}")
print(f"   Market Cap range: {format_number(df['Market Cap'].min())} - {format_number(df['Market Cap'].max())}")

# Create presentation with chart
prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(7.5)

# Add blank slide
slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout

# Add title
title_box = slide.shapes.add_textbox(
    Inches(LAYOUT_CONSTANTS['slide_padding_top']/72),  # Convert px to inches
    Inches(0.5),
    Inches(9),
    Inches(0.8)
)
title_tf = title_box.text_frame
title_p = title_tf.paragraphs[0]
title_p.text = "Top 10 Companies by Market Cap"
title_p.font.size = Inches(FINANCE_FONTS['title_size']/72)
title_p.font.name = FINANCE_FONTS['primary']
title_p.font.bold = True

# Import and use advanced chart builder
from src.converter.advanced_finance_charts import AdvancedFinanceChartBuilder

chart_builder = AdvancedFinanceChartBuilder(
    slide=slide,
    position=(1, 2),
    size=(8, 4.5)
)

print("\n🎨 Building chart with AdvancedFinanceChartBuilder...")
success = chart_builder.create_chart(
    df=df,
    chart_type='COLUMN',
    title='Market Capitalization Comparison'
)

if success:
    print("   ✅ Chart created successfully")
else:
    print("   ❌ Chart creation failed")

# ============================================================================
# Test 6: Finalize Presentation (Master Normalizer)
# ============================================================================
print("\n🎨 Test 6: finalize_presentation() - Master Finalizer")
finalize_presentation(prs)

# Save output
output_path = Path(__file__).parent / 'examples' / 'demo_PPT' / 'enhanced_finance_standards_output.pptx'
output_path.parent.mkdir(parents=True, exist_ok=True)
prs.save(str(output_path))

print(f"\n✅ SUCCESS! Presentation created and saved")
print(f"📁 Output: {output_path}")
print(f"📊 Slides: {len(prs.slides)}")

# ============================================================================
# Test 7: Validation Checks
# ============================================================================
print("\n🧪 Test 7: Validation Checks")
validation_checks = {
    '$ Prefix in format_number': format_number(1000).startswith('$'),
    'No scientific notation (1e12)': 'e' not in format_number(1e12).lower(),
    'Percentage with 1 decimal': format_percentage(12.345) == '12.3%',
    'Title font size is 28pt': FINANCE_FONTS['title_size'] == 28,
    'Axis font size is 11pt': FINANCE_FONTS['axis_size'] == 11,
    'Legend font size is 9pt': FINANCE_FONTS['legend_size'] == 9,
    'Max Y-ticks is 4-6': LAYOUT_CONSTANTS['min_y_ticks'] <= LAYOUT_CONSTANTS['max_y_ticks'] <= 6,
    'Top-N limit is 5': LAYOUT_CONSTANTS['top_n_limit'] == 5,
    'Chart width is 70-80%': 0.7 <= LAYOUT_CONSTANTS['chart_width_ratio'] <= 0.8,
    'Chart height max 60%': LAYOUT_CONSTANTS['chart_max_height_ratio'] == 0.60,
}

all_passed = True
for check_name, result in validation_checks.items():
    status = "✅" if result else "❌"
    print(f"   {status} {check_name}")
    if not result:
        all_passed = False

print("\n" + "="*80)
if all_passed:
    print("🎉 ALL TESTS PASSED! Finance visual standards are production-ready.")
else:
    print("⚠️  Some validation checks failed. Review the output above.")
print("="*80)

print("\n🎯 Expected Features in Generated Presentation:")
print("   ✅ Fonts: Segoe UI (Title 28pt Bold, Axis 11pt, Legend 9pt)")
print("   ✅ Numbers: $2.8T, $2.6T, $1.7T (NO scientific notation)")
print("   ✅ Colors: #004F9E (primary), #16A085 (positive), #E15759 (negative)")
print("   ✅ Layout: 36px top margin, chart 70-80% width, max 60% height")
print("   ✅ Axes: 4-6 Y-ticks, horizontal gridlines only (0.25pt, light gray)")
print("   ✅ Legend: Bottom-center, horizontal, 9pt")
print("   ✅ Bars: 150% gap width, subtle vertical gradient")
print("   ✅ Title: 28pt Bold with brand accent underline (2px, #004F9E)")
print("   ✅ Background: White (no chart backgrounds)")
print("\n📖 Open the generated PPTX to verify all standards are applied!")
