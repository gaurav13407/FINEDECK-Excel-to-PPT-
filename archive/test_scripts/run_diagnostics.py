# DIAGNOSTIC CHECKLIST - Run this to verify your setup

print("="*70)
print(" CHART SYSTEM DIAGNOSTIC CHECKLIST")
print("="*70)

# Check 1: SimpleFinanceChartBuilder exists
print("\n1️⃣  Checking SimpleFinanceChartBuilder file exists...")
import os
chart_file = r"c:\Users\gaura\OneDrive\Desktop\Big Projects\FinDeck(Excel to PPT Project)\src\converter\simple_finance_charts.py"
if os.path.exists(chart_file):
    print(f"   ✅ File exists: {chart_file}")
else:
    print(f"   ❌ File NOT FOUND: {chart_file}")

# Check 2: Can import SimpleFinanceChartBuilder
print("\n2️⃣  Checking if SimpleFinanceChartBuilder can be imported...")
try:
    from src.converter.simple_finance_charts import SimpleFinanceChartBuilder
    print("   ✅ Import successful")
except Exception as e:
    print(f"   ❌ Import failed: {e}")

# Check 3: EnhancedProfessionalBuilder imports it
print("\n3️⃣  Checking EnhancedProfessionalBuilder uses SimpleFinanceChartBuilder...")
try:
    with open(r"c:\Users\gaura\OneDrive\Desktop\Big Projects\FinDeck(Excel to PPT Project)\src\converter\enhanced_professional_builder.py", 'r', encoding='utf-8') as f:
        content = f.read()
        if 'from src.converter.simple_finance_charts import SimpleFinanceChartBuilder' in content:
            print("   ✅ Import statement found")
        else:
            print("   ❌ Import statement NOT found!")
        
        if 'self.chart_builder = SimpleFinanceChartBuilder' in content:
            print("   ✅ Initialization found")
        else:
            print("   ❌ Initialization NOT found!")
        
        if 'self.chart_builder.create_chart' in content:
            print("   ✅ create_chart() calls found")
        else:
            print("   ❌ create_chart() calls NOT found!")
except Exception as e:
    print(f"   ❌ Error reading file: {e}")

# Check 4: No fallback methods remain
print("\n4️⃣  Checking for removed fallback methods...")
try:
    with open(r"c:\Users\gaura\OneDrive\Desktop\Big Projects\FinDeck(Excel to PPT Project)\src\converter\enhanced_professional_builder.py", 'r', encoding='utf-8') as f:
        content = f.read()
        
        bad_methods = [
            '_add_insights_chart_fallback',
            '_add_trend_chart_fallback',
            'self.chart_analyzer',
            'self.advanced_chart_builder.create_chart'
        ]
        
        for method in bad_methods:
            if method in content:
                print(f"   ❌ Found old method: {method}")
            else:
                print(f"   ✅ {method} removed")
except Exception as e:
    print(f"   ❌ Error: {e}")

# Check 5: Test chart creation
print("\n5️⃣  Testing chart creation...")
try:
    import pandas as pd
    from pptx import Presentation
    from pptx.util import Inches
    from src.converter.simple_finance_charts import SimpleFinanceChartBuilder
    
    # Create test data
    test_df = pd.DataFrame({
        'Sector': ['Tech', 'Health', 'Finance'],
        'Allocation': [40, 35, 25]
    })
    
    # Create builder
    builder = SimpleFinanceChartBuilder()
    
    # Create presentation and slide
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    
    # Try to create chart
    chart = builder.create_chart(
        slide, test_df,
        left=Inches(1.0),
        top=Inches(2.0),
        width=Inches(4.0),
        height=Inches(3.0),
        title="Test Chart"
    )
    
    if chart:
        print("   ✅ Chart created successfully")
        
        # Check position
        for shape in slide.shapes:
            if hasattr(shape, 'chart'):
                if shape.left < 100000000:  # Less than 100 million
                    print(f"   ✅ Position correct: {shape.left} EMUs ({shape.left/914400:.2f} inches)")
                else:
                    print(f"   ❌ Position WRONG: {shape.left} EMUs (TRILLIONS!)")
    else:
        print("   ❌ Chart creation returned None")
        
except Exception as e:
    print(f"   ❌ Error: {e}")
    import traceback
    traceback.print_exc()

# Check 6: Check __pycache__ (might have old code)
print("\n6️⃣  Checking for cached Python files...")
cache_dirs = []
for root, dirs, files in os.walk(r"c:\Users\gaura\OneDrive\Desktop\Big Projects\FinDeck(Excel to PPT Project)\src\converter"):
    if '__pycache__' in dirs:
        cache_dir = os.path.join(root, '__pycache__')
        cache_dirs.append(cache_dir)

if cache_dirs:
    print(f"   ⚠️  Found {len(cache_dirs)} __pycache__ directories:")
    for cache_dir in cache_dirs:
        print(f"      - {cache_dir}")
    print("   💡 These should be deleted before restart!")
else:
    print("   ✅ No __pycache__ directories found")

# Check 7: Test backend converter flow
print("\n7️⃣  Testing backend converter flow...")
try:
    from src.converter.excel_to_ppt_converter import ExcelToPPTConverter
    
    converter = ExcelToPPTConverter(
        user_tier='ai_pro',
        user_id='test',
        user_metadata={'name': 'Test', 'email': 'test@test.com', 'company': 'Test'},
        use_finance_charts=True
    )
    print("   ✅ ExcelToPPTConverter created successfully")
    print(f"      Tier: {converter.user_tier}")
    print(f"      Finance charts: {converter.use_finance_charts}")
except Exception as e:
    print(f"   ❌ Error: {e}")

print("\n" + "="*70)
print(" DIAGNOSTIC COMPLETE")
print("="*70)

print("\n📝 SUMMARY:")
print("   If ALL checks show ✅:")
print("      → Your code is correct")
print("      → Problem is backend cache")
print("      → SOLUTION: Kill Python, delete __pycache__, restart")
print("")
print("   If ANY check shows ❌:")
print("      → Code has issues")
print("      → Check the specific failing item above")
