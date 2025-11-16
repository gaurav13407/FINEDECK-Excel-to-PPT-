"""
COMPLETE PIPELINE DIAGNOSTIC
Trace the entire flow from browser → backend → converter → charts
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

print("\n" + "="*80)
print("🔍 COMPLETE BACKEND PIPELINE DIAGNOSTIC")
print("="*80)

print("\n1️⃣ CHECKING ENHANCED CHART BUILDER CONNECTION")
print("-" * 80)

# Check if enhanced_professional_builder imports enhanced charts
from src.converter.enhanced_professional_builder import EnhancedProfessionalBuilder

print("✅ EnhancedProfessionalBuilder imported successfully")

# Check if it has the enhanced_chart_builder attribute
builder = EnhancedProfessionalBuilder(
    ai_service=None,
    user_metadata={'name': 'Test', 'email': 'test@test.com', 'company': 'Test'},
    user_tier='ai_pro',
    use_finance_charts=False
)

print(f"✅ Builder created for tier: ai_pro")
print(f"   - Has enhanced_chart_builder attribute: {hasattr(builder, 'enhanced_chart_builder')}")
print(f"   - Has visual_enhancer attribute: {hasattr(builder, 'visual_enhancer')}")

# Check which chart methods exist
chart_methods = [
    '_add_insights_chart',
    '_add_sector_chart', 
    '_add_trend_chart',
    '_add_insights_chart_fallback',
    '_add_trend_chart_fallback'
]

for method in chart_methods:
    exists = hasattr(builder, method)
    print(f"   - Method {method}: {'✅' if exists else '❌'}")

print("\n2️⃣ CHECKING ENHANCED CHART BUILDER METHODS")
print("-" * 80)

from src.converter.enhanced_charts import EnhancedChartBuilder

# Create with dummy template colors
template_colors = {
    'chart_colors': [
        (123, 31, 162),   # Purple
        (171, 71, 188),
        (186, 104, 200)
    ]
}

ecb = EnhancedChartBuilder(template_colors)
print("✅ EnhancedChartBuilder created")

chart_builder_methods = [
    'detect_chart_type',
    'create_pie_chart',
    'create_bar_chart',
    'create_line_chart',
    'create_column_chart',
    'auto_create_chart'
]

for method in chart_builder_methods:
    exists = hasattr(ecb, method)
    print(f"   - Method {method}: {'✅' if exists else '❌'}")

print("\n3️⃣ TESTING CHART TYPE DETECTION")
print("-" * 80)

import pandas as pd

# Test data scenarios
test_scenarios = [
    ("Portfolio Allocation", pd.DataFrame({
        'Sector': ['Tech', 'Finance', 'Healthcare', 'Energy', 'Consumer'],
        'Allocation %': [30, 25, 20, 15, 10]
    })),
    ("Revenue Trend", pd.DataFrame({
        'Quarter': ['Q1 2024', 'Q2 2024', 'Q3 2024', 'Q4 2024'],
        'Revenue': [100, 120, 140, 160]
    })),
    ("Top 5 Products", pd.DataFrame({
        'Product': ['A', 'B', 'C', 'D', 'E'],
        'Sales': [500, 450, 400, 350, 300]
    }))
]

for name, df in test_scenarios:
    detected = ecb.detect_chart_type(df)
    print(f"   {name}:")
    print(f"      Columns: {list(df.columns)}")
    print(f"      Detected type: {detected.upper()}")

print("\n4️⃣ CHECKING BACKEND API ENDPOINT")
print("-" * 80)

# Check if tiered_conversions.py calls convert_professional
with open('src/backend/app/api/v1/endpoints/tiered_conversions.py', 'r') as f:
    content = f.read()
    has_convert_professional = 'convert_professional' in content
    has_template_name = 'template_name' in content
    has_enhanced_builder = 'EnhancedProfessionalBuilder' in content
    
    print(f"   - Calls convert_professional: {'✅' if has_convert_professional else '❌'}")
    print(f"   - Passes template_name: {'✅' if has_template_name else '❌'}")

print("\n5️⃣ CHECKING EXCEL_TO_PPT_CONVERTER")
print("-" * 80)

# Check if converter imports and uses EnhancedProfessionalBuilder
with open('src/converter/excel_to_ppt_converter.py', 'r') as f:
    content = f.read()
    imports_enhanced = 'from src.converter.enhanced_professional_builder import EnhancedProfessionalBuilder' in content
    creates_enhanced = 'EnhancedProfessionalBuilder(' in content
    
    print(f"   - Imports EnhancedProfessionalBuilder: {'✅' if imports_enhanced else '❌'}")
    print(f"   - Creates EnhancedProfessionalBuilder instance: {'✅' if creates_enhanced else '❌'}")

print("\n6️⃣ CHECKING BUILD_PRESENTATION METHOD")
print("-" * 80)

# Check if build_presentation initializes enhancers
with open('src/converter/enhanced_professional_builder.py', 'r') as f:
    content = f.read()
    
    # Check for initialization
    init_visual = 'self.visual_enhancer = VisualEnhancer(self.template_colors)' in content
    init_chart = 'self.enhanced_chart_builder = EnhancedChartBuilder(self.template_colors)' in content
    
    # Check for usage in chart methods
    uses_enhanced_insights = 'self.enhanced_chart_builder.auto_create_chart' in content
    
    print(f"   - Initializes VisualEnhancer: {'✅' if init_visual else '❌'}")
    print(f"   - Initializes EnhancedChartBuilder: {'✅' if init_chart else '❌'}")
    print(f"   - Uses enhanced_chart_builder.auto_create_chart: {'✅' if uses_enhanced_insights else '❌'}")
    
    # Count occurrences
    occurrences = content.count('self.enhanced_chart_builder.auto_create_chart')
    print(f"   - Number of times enhanced charts are called: {occurrences}")

print("\n7️⃣ PIPELINE FLOW SUMMARY")
print("-" * 80)

print("""
EXPECTED FLOW:
1. Browser sends: template_name='royal_purple', Excel file
2. Backend API (tiered_conversions.py): 
   → Receives request
   → Creates ExcelToPPTConverter(user_tier='ai_pro')
   → Calls converter.convert_professional(template_name='royal_purple')
3. ExcelToPPTConverter.convert_professional():
   → Validates template is in allowed list
   → Loads template JSON (royal_purple.json)
   → Creates EnhancedProfessionalBuilder
   → Calls slide_builder.build_presentation(template=template)
4. EnhancedProfessionalBuilder.build_presentation():
   → Initializes self.visual_enhancer = VisualEnhancer(template_colors)
   → Initializes self.enhanced_chart_builder = EnhancedChartBuilder(template_colors)
   → Creates slides with charts:
      - Data Insights slide → calls self._add_insights_chart()
      - Sector Distribution slide → calls enhanced_chart_builder.auto_create_chart()
      - Trend Analysis slide → calls self._add_trend_chart()
5. EnhancedChartBuilder.auto_create_chart():
   → Detects chart type (PIE/BAR/LINE/COLUMN)
   → Creates appropriate chart with purple template colors
   → Returns chart object
6. PPT saved and returned to browser
""")

print("\n" + "="*80)
print("🎯 DIAGNOSIS")
print("="*80)

print("""
BASED ON YOUR ISSUE:
1. ✅ Charts ARE being created (verified in downloaded files)
2. ✅ Enhanced chart system IS connected
3. ❌ You're getting BAR charts but want different types
4. ❌ Template colors might be wrong (corporate_blue instead of royal_purple)
5. ❌ PPT might be "broken" when downloading

MOST LIKELY ISSUES:
1. Template validation failing → Falls back to corporate_blue
2. Chart type auto-detection creating BAR (horizontal) instead of PIE/COLUMN
3. PPT corruption during download (check file size)

TO FIX:
1. Check backend logs for template validation messages
2. Adjust chart auto-detection thresholds
3. Verify template_name is being passed correctly
4. Check for download errors in browser
""")

print("\n" + "="*80)
