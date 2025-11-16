"""
Direct test of EnhancedProfessionalBuilder with multiple chart showcase
"""

import pandas as pd
import sys
import os
from pptx import Presentation

sys.path.insert(0, os.path.abspath('.'))

from src.converter.enhanced_professional_builder import EnhancedProfessionalBuilder

def test_chart_showcase():
    """Test the new multiple chart showcase feature"""
    
    print("\n" + "="*80)
    print("🎨 TESTING MULTIPLE CHART SHOWCASE FEATURE")
    print("="*80 + "\n")
    
    # Create test data
    print("📊 Creating test data...")
    
    # Quarterly data (for LINE, AREA, COLUMN, STACKED)
    quarters_data = pd.DataFrame({
        'Quarter': ['Q1 2024', 'Q2 2024', 'Q3 2024', 'Q4 2024', 'Q1 2025', 'Q2 2025'],
        'Revenue': [125000, 142000, 158000, 189000, 215000, 248000],
        'Profit': [35000, 42000, 48000, 62000, 78000, 92000],
        'Expenses': [90000, 100000, 110000, 127000, 137000, 156000],
        'Product A': [45000, 52000, 58000, 68000, 78000, 92000],
        'Product B': [38000, 42000, 48000, 55000, 62000, 70000],
        'Product C': [25000, 28000, 32000, 38000, 45000, 52000]
    })
    
    # Portfolio allocation (for DONUT)
    allocation_data = pd.DataFrame({
        'Sector': ['Technology', 'Healthcare', 'Finance', 'Energy', 'Consumer'],
        'Allocation': [35, 25, 20, 12, 8]
    })
    
    # Stock data (for CANDLESTICK)
    stock_data = pd.DataFrame({
        'Date': ['2025-01-01', '2025-01-02', '2025-01-03', '2025-01-04', '2025-01-05'],
        'Open': [150.2, 152.5, 151.8, 153.2, 154.5],
        'High': [153.8, 154.2, 154.5, 155.8, 156.2],
        'Low': [149.5, 151.2, 150.8, 152.5, 153.2],
        'Close': [152.5, 151.8, 153.2, 154.5, 155.8]
    })
    
    # Combine data
    all_data = pd.concat([quarters_data, allocation_data, stock_data], ignore_index=True)
    
    sheets_data = [
        ('Quarterly Performance', quarters_data),
        ('Portfolio Allocation', allocation_data),
        ('Stock Prices', stock_data)
    ]
    
    print(f"   ✅ Created {len(sheets_data)} sheets with diverse data\n")
    
    # Create presentation
    print("🎨 Building presentation with EnhancedProfessionalBuilder...")
    prs = Presentation()
    prs.slide_width = 9144000  # 10 inches
    prs.slide_height = 6858000  # 7.5 inches
    
    # Initialize builder
    builder = EnhancedProfessionalBuilder(
        ai_service=None,
        user_metadata=None,
        user_tier='pro',
        use_finance_charts=True
    )
    
    # Build presentation
    results = builder.build_presentation(
        prs=prs,
        sheets_data=sheets_data,
        project_name="Multi-Chart Financial Analysis",
        template=None
    )
    
    # Save
    output_path = "examples/demo_PPT/chart_showcase_direct_test.pptx"
    prs.save(output_path)
    
    print("\n" + "="*80)
    print("🎉 PRESENTATION CREATED!")
    print("="*80)
    print(f"📁 Output: {output_path}")
    print(f"📊 Total slides: {results.get('slides_created', 'N/A')}")
    print(f"📋 Slide names:")
    for name in results.get('slide_names', []):
        print(f"   • {name}")
    print("\n✅ Look for 'Chart Showcase' slides with different chart types!")
    print("="*80 + "\n")

if __name__ == "__main__":
    test_chart_showcase()
