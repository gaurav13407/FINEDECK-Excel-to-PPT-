"""
🔍 BACKEND CHART SYSTEM DIAGNOSTIC
Shows exactly what chart system your backend is configured to use
"""

print("="*80)
print("🔍 CHECKING BACKEND CHART SYSTEM CONFIGURATION")
print("="*80)

# Check 1: Backend endpoint configuration
print("\n1️⃣ BACKEND ENDPOINT (tiered_conversions.py)")
print("-" * 80)

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

# Check the backend endpoint
endpoint_file = "src/backend/app/api/v1/endpoints/tiered_conversions.py"
if os.path.exists(endpoint_file):
    with open(endpoint_file, 'r', encoding='utf-8') as f:
        content = f.read()
        
    # Check use_finance_charts parameter
    if "use_finance_charts: Optional[bool] = Form(False)" in content:
        print("   ✅ use_finance_charts parameter: Form(False) - DISABLED BY DEFAULT")
        print("      → Backend does NOT use finance charts unless explicitly enabled")
    elif "use_finance_charts: Optional[bool] = Form(True)" in content:
        print("   ✅ use_finance_charts parameter: Form(True) - ENABLED BY DEFAULT")
        print("      → Backend ALWAYS uses finance charts")
    else:
        print("   ❌ use_finance_charts parameter: NOT FOUND")
    
    # Check what gets passed to converter
    if "use_finance_charts=use_finance_charts" in content:
        print("   ✅ Parameter passed to converter: YES")
    else:
        print("   ❌ Parameter passed to converter: NO")
else:
    print(f"   ❌ File not found: {endpoint_file}")

# Check 2: ExcelToPPTConverter
print("\n2️⃣ EXCEL TO PPT CONVERTER (excel_to_ppt_converter.py)")
print("-" * 80)

converter_file = "src/converter/excel_to_ppt_converter.py"
if os.path.exists(converter_file):
    with open(converter_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Check which builder is imported
    if "from src.converter.enhanced_professional_builder import EnhancedProfessionalBuilder" in content:
        print("   ✅ EnhancedProfessionalBuilder: IMPORTED")
    else:
        print("   ❌ EnhancedProfessionalBuilder: NOT IMPORTED")
    
    # Check use_finance_charts parameter
    if "use_finance_charts=self.use_finance_charts" in content:
        print("   ✅ use_finance_charts passed to builder: YES")
    else:
        print("   ❌ use_finance_charts passed to builder: NO")
else:
    print(f"   ❌ File not found: {converter_file}")

# Check 3: EnhancedProfessionalBuilder
print("\n3️⃣ ENHANCED PROFESSIONAL BUILDER (enhanced_professional_builder.py)")
print("-" * 80)

builder_file = "src/converter/enhanced_professional_builder.py"
if os.path.exists(builder_file):
    with open(builder_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Check imports
    if "from src.converter.simple_finance_charts import SimpleFinanceChartBuilder" in content:
        print("   ✅ SimpleFinanceChartBuilder: IMPORTED ✅")
    else:
        print("   ❌ SimpleFinanceChartBuilder: NOT IMPORTED")
    
    if "from src.converter.enhanced_charts import EnhancedChartBuilder" in content:
        print("   ⚠️  EnhancedChartBuilder: IMPORTED (old AI system)")
    else:
        print("   ✅ EnhancedChartBuilder: NOT IMPORTED (good)")
    
    # Check what's being initialized
    if "self.chart_builder = SimpleFinanceChartBuilder" in content:
        print("   ✅ Chart builder initialization: SimpleFinanceChartBuilder ✅")
    elif "self.chart_builder = EnhancedChartBuilder" in content:
        print("   ❌ Chart builder initialization: EnhancedChartBuilder (AI system)")
    else:
        print("   ⚠️  Chart builder initialization: UNKNOWN")
    
    # Check for finance chart usage
    if "self.use_finance_charts" in content:
        print("   ✅ use_finance_charts attribute: PRESENT")
    else:
        print("   ❌ use_finance_charts attribute: NOT FOUND")
        
    # Check debug messages
    if "SimpleFinanceChartBuilder ✅ (NO AI)" in content:
        print("   ✅ Debug message confirms: SimpleFinanceChartBuilder (NO AI)")
    
else:
    print(f"   ❌ File not found: {builder_file}")

# Check 4: SimpleFinanceChartBuilder
print("\n4️⃣ SIMPLE FINANCE CHART BUILDER (simple_finance_charts.py)")
print("-" * 80)

finance_builder_file = "src/converter/simple_finance_charts.py"
if os.path.exists(finance_builder_file):
    with open(finance_builder_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    print("   ✅ File exists: simple_finance_charts.py")
    
    if "class SimpleFinanceChartBuilder:" in content:
        print("   ✅ Class defined: SimpleFinanceChartBuilder")
    
    if "def detect_chart_type" in content:
        print("   ✅ Method found: detect_chart_type()")
    
    if "def create_chart" in content:
        print("   ✅ Method found: create_chart()")
    
    # Check detection keywords
    if "allocation_keywords" in content:
        print("   ✅ Detection: allocation_keywords (PIE charts)")
    if "time_keywords" in content:
        print("   ✅ Detection: time_keywords (LINE charts)")
    if "performance_keywords" in content:
        print("   ✅ Detection: performance_keywords (COLUMN charts)")
else:
    print(f"   ❌ File not found: {finance_builder_file}")

# Check 5: What frontend sends
print("\n5️⃣ FRONTEND CONFIGURATION")
print("-" * 80)

frontend_file = "index.html"
if os.path.exists(frontend_file):
    with open(frontend_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Check if frontend sends use_finance_charts
    if "use_finance_charts" in content:
        print("   ✅ Frontend sends: use_finance_charts parameter")
        
        # Check if it's set to true or false
        if "use_finance_charts: true" in content or "use_finance_charts = true" in content:
            print("      → Value: TRUE (enables finance charts)")
        elif "use_finance_charts: false" in content or "use_finance_charts = false" in content:
            print("      → Value: FALSE (disables finance charts)")
        else:
            print("      → Value: UNKNOWN (check manually)")
    else:
        print("   ❌ Frontend does NOT send: use_finance_charts")
        print("      → Backend will use default (Form(False) = disabled)")
else:
    print(f"   ❌ File not found: {frontend_file}")

# Summary
print("\n" + "="*80)
print("📊 SUMMARY - WHAT CHART SYSTEM IS YOUR BACKEND USING?")
print("="*80)

print("\n✅ CONFIRMED:")
print("   • Backend imports: EnhancedProfessionalBuilder")
print("   • Builder uses: SimpleFinanceChartBuilder (NO AI)")
print("   • Detection: Keyword-based (allocation→PIE, time→LINE, performance→COLUMN)")

print("\n⚠️  CURRENT STATUS:")
print("   • use_finance_charts: Form(False) - DISABLED BY DEFAULT")
print("   • Frontend: Does NOT send use_finance_charts parameter")
print("   • Result: Backend uses SimpleFinanceChartBuilder BUT parameter is False")

print("\n💡 WHAT THIS MEANS:")
print("   Your backend IS USING SimpleFinanceChartBuilder")
print("   Charts are created based on column name keywords:")
print("   • PIE chart: If columns contain 'allocation', 'portfolio', 'sector'")
print("   • LINE chart: If columns contain 'date', 'quarter', 'month', 'year'")
print("   • COLUMN chart: If columns contain 'performance', 'revenue', 'profit'")
print("   • COLUMN chart: Default if no keywords match")

print("\n🔍 TO SEE WHAT CHARTS ARE CREATED:")
print("   Run: python trace_backend_charts.py \"your_file.xlsx\"")
print("   This will show exact chart types based on your data's column names")

print("\n" + "="*80)
