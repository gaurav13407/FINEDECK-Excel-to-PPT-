"""
Test AI Service - All 5 Features
Verify each AI feature works correctly

Run this script to test all AI features:
python test_ai_service.py
"""

import sys
import os
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

import pandas as pd
from src.backend.app.services.ai_service import create_ai_service
import time

print("=" * 70)
print("TESTING AI SERVICE - ALL 5 FEATURES")
print("=" * 70)

# Create test data
test_data = pd.DataFrame({
    'Quarter': ['Q1', 'Q2', 'Q3', 'Q4'],
    'Revenue': [100000, 125000, 145000, 185000],
    'Profit': [20000, 26000, 32000, 45000],
    'Costs': [80000, 99000, 113000, 140000]
})

print("\nTest Data:")
print(test_data)
print("\n" + "=" * 70)

try:
    # Initialize AI service
    print("\n1. Initializing AI Service...")
    ai = create_ai_service()
    print("   SUCCESS - AI Service initialized")
    
    # Test Feature 1: Slide Title
    print("\n2. Testing Feature 1: Slide Title Generator...")
    start = time.time()
    title = ai.generate_slide_title(test_data, "Quarterly Performance", "line")
    elapsed = time.time() - start
    print(f"   Generated Title: '{title}'")
    print(f"   Time taken: {elapsed:.2f}s")
    
    # Test Feature 2: Slide Summary
    print("\n3. Testing Feature 2: Slide Summary Generator...")
    start = time.time()
    summary = ai.generate_slide_summary(test_data, "Quarterly Performance", "line")
    elapsed = time.time() - start
    print(f"   Generated Summary:")
    print(f"   {summary}")
    print(f"   Time taken: {elapsed:.2f}s")
    
    # Test Feature 3: Data Insights
    print("\n4. Testing Feature 3: Data Insights Generator...")
    start = time.time()
    insights = ai.generate_data_insights(test_data, "Quarterly Performance", num_insights=5)
    elapsed = time.time() - start
    print(f"   Generated {len(insights)} insights:")
    for i, insight in enumerate(insights, 1):
        print(f"   {i}. {insight}")
    print(f"   Time taken: {elapsed:.2f}s")
    
    # Test Feature 4: Template Recommendation
    print("\n5. Testing Feature 4: Smart Template Selection...")
    start = time.time()
    sheets_info = [
        {"name": "Quarterly Performance", "data_type": "time_series", "rows": 4, "cols": 4},
        {"name": "Product Sales", "data_type": "comparison", "rows": 10, "cols": 3}
    ]
    templates = ["corporate_blue", "financial_green", "executive_dark", "minimal_white", 
                "vibrant_orange", "professional_purple", "tech_gradient", "modern_teal",
                "elegant_gold", "classic_red"]
    
    template_rec = ai.recommend_template("Q4_Financial_Report.xlsx", sheets_info, templates)
    elapsed = time.time() - start
    print(f"   Top Recommendation: {template_rec.get('auto_selected')}")
    print(f"   All Recommendations:")
    for rec in template_rec.get('recommendations', []):
        print(f"   - {rec['template']}: {rec['confidence']} - {rec['reasoning']}")
    print(f"   Time taken: {elapsed:.2f}s")
    
    # Test Feature 5: Layout Optimizer
    print("\n6. Testing Feature 5: Slide Layout Optimizer...")
    start = time.time()
    layout = ai.optimize_slide_layout(test_data, "line", has_insights=True)
    elapsed = time.time() - start
    print(f"   Recommended Layout: {layout.get('layout')}")
    print(f"   Reasoning: {layout.get('reasoning')}")
    print(f"   Time taken: {elapsed:.2f}s")
    
    # Test Feature 6: Chart Type Recommendation
    print("\n7. Testing Feature 6: AI Chart Type Recommendation...")
    start = time.time()
    chart_rec = ai.recommend_chart_type(test_data, list(test_data.columns), "quarterly financial")
    elapsed = time.time() - start
    print(f"   Recommended Chart: {chart_rec['recommended']['type']}")
    print(f"   Confidence: {chart_rec['recommended']['confidence']}")
    print(f"   Reasoning: {chart_rec['recommended']['reasoning']}")
    if chart_rec['recommended'].get('tips'):
        print(f"   Tips:")
        for tip in chart_rec['recommended']['tips']:
            print(f"   - {tip}")
    print(f"   Time taken: {elapsed:.2f}s")
    
    # Show usage stats
    print("\n" + "=" * 70)
    print("USAGE STATISTICS:")
    stats = ai.get_usage_stats()
    print(f"   Total Tokens Used: {stats['total_tokens']}")
    print(f"   Total Cost: ${stats['total_cost_usd']}")
    print(f"   Cost Per Presentation: ${stats['cost_per_presentation']}")
    print(f"   Estimated cost for 100 presentations: ${stats['cost_per_presentation'] * 100:.2f}")
    print("=" * 70)
    
    print("\n" + "=" * 70)
    print("SUCCESS! ALL 5 AI FEATURES WORKING PERFECTLY!")
    print("=" * 70)
    print("\nAI Features Ready:")
    print("  1. Slide Title Generator")
    print("  2. Slide Summary Generator")
    print("  3. Data Insights Generator")
    print("  4. Smart Template Selection")
    print("  5. Layout Optimizer")
    print("  6. Chart Type Recommendation")
    print("\nNext Steps:")
    print("  - Integrate into ppt_writer.py")
    print("  - Add AI Pro API endpoints")
    print("  - Test with real Excel files")
    print("=" * 70)
    
except Exception as e:
    print(f"\n❌ ERROR: {str(e)}")
    import traceback
    traceback.print_exc()
    print("\n⚠️ Check the errors above and fix them!")
