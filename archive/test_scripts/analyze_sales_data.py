import sys
sys.path.insert(0, r'c:\Users\gaura\OneDrive\Desktop\Big Projects\FinDeck(Excel to PPT Project)\src\backend\app')

from services.data_intelligence import DataIntelligenceEngine
import json

# Analyze the DV+Sales+Data.xlsx file
engine = DataIntelligenceEngine()
result = engine.analyze_file(r'c:\Users\gaura\OneDrive\Desktop\Big Projects\FinDeck(Excel to PPT Project)\DV+Sales+Data.xlsx')

# Save analysis to JSON
with open('sales_data_analysis.json', 'w') as f:
    json.dump(result, f, indent=2)

print("✅ Analysis Complete!")
print("\n" + "="*80)
print("EXECUTIVE SUMMARY")
print("="*80)
print(json.dumps(result['executive_summary'], indent=2))

print("\n" + "="*80)
print("KEY METRICS")
print("="*80)
for metric in result['key_metrics'][:5]:  # Show top 5
    print(f"\n📊 {metric['metric']}")
    print(f"   Value: {metric['value']}")
    print(f"   Insight: {metric['insight']}")

print("\n" + "="*80)
print("HIERARCHY ANALYSIS")
print("="*80)
print(json.dumps(result['hierarchy_analysis'], indent=2))

print("\n✅ Full analysis saved to: sales_data_analysis.json")
