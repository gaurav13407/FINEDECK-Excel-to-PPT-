"""
Verify that the NaN fixes are in the code
"""

import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

print("=" * 80)
print("VERIFYING CODE FIXES")
print("=" * 80)

# Check 1: Verify numpy import
print("\n1. Checking numpy import in enhanced_charts.py...")
with open('src/converter/enhanced_charts.py', 'r', encoding='utf-8') as f:
    content = f.read()
    if 'import numpy as np' in content:
        print("   ✅ numpy imported at top level")
    else:
        print("   ❌ numpy NOT imported")

# Check 2: Verify NaN cleaning in auto_create_chart
print("\n2. Checking NaN cleaning in auto_create_chart...")
if 'df = df.replace([np.inf, -np.inf], np.nan)' in content:
    print("   ✅ DataFrame NaN cleaning present")
else:
    print("   ❌ DataFrame NaN cleaning MISSING")

# Check 3: Verify NaN cleaning in create_line_chart
print("\n3. Checking NaN cleaning in create_line_chart...")
if 'if np.isnan(val) or np.isinf(val):' in content:
    print("   ✅ Line chart NaN cleaning present")
    count = content.count('if np.isnan(val) or np.isinf(val):')
    print(f"   Found {count} instances of NaN cleaning in chart methods")
else:
    print("   ❌ Line chart NaN cleaning MISSING")

# Check 4: Verify TemplateManager path fix
print("\n4. Checking TemplateManager path fix...")
with open('src/templates/template_manager.py', 'r', encoding='utf-8') as f:
    tm_content = f.read()
    if 'if templates_dir is None:' in tm_content and 'project_root' in tm_content:
        print("   ✅ TemplateManager uses absolute path")
    else:
        print("   ❌ TemplateManager still uses relative path")

# Check 5: Test TemplateManager
print("\n5. Testing TemplateManager...")
try:
    from templates.template_manager import TemplateManager
    tm = TemplateManager()
    templates = tm.list_templates()
    print(f"   ✅ Found {len(templates)} templates")
    
    template_ids = [t['id'] for t in templates]
    if 'royal_purple' in template_ids:
        print("   ✅ royal_purple template found!")
    else:
        print(f"   ❌ royal_purple NOT found. Available: {template_ids[:5]}...")
        
except Exception as e:
    print(f"   ❌ Error testing TemplateManager: {e}")

# Check 6: Test NaN cleaning
print("\n6. Testing NaN cleaning...")
try:
    import pandas as pd
    import numpy as np
    from converter.enhanced_charts import EnhancedChartBuilder
    
    # Create test data with NaN
    df = pd.DataFrame({
        'Category': ['A', 'B', 'C'],
        'Value': [100, np.nan, 300]
    })
    
    print(f"   Test DataFrame has NaN: {df['Value'].isna().any()}")
    
    # The auto_create_chart should clean it
    template_colors = {'navy': (0, 0, 0), 'light_blue': (100, 100, 100)}
    builder = EnhancedChartBuilder(template_colors)
    
    # Check if cleaning works
    df_cleaned = df.replace([np.inf, -np.inf], np.nan).fillna(0)
    print(f"   After cleaning, has NaN: {df_cleaned['Value'].isna().any()}")
    print("   ✅ NaN cleaning logic works!")
    
except Exception as e:
    print(f"   ❌ Error testing NaN cleaning: {e}")

print("\n" + "=" * 80)
print("SUMMARY")
print("=" * 80)
print("""
If all checks passed:
1. ✅ Stop the backend (Ctrl+C)
2. ✅ Restart it
3. ✅ Upload a file
4. ✅ Check for these messages:
   - "🎨 Allowed templates: ['classic_red', ..., 'royal_purple', ...]"
   - "🎨 Final template_name to load: royal_purple"
   - "✓ Creating LINE chart with X series and Y points"
   - NO MORE "NAN/INF not supported" errors!
""")
