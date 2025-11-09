"""
Professional PPT Enhancements Demo
===================================
This script demonstrates all the new visual enhancements:
1. Enhanced cover page with branding
2. Multiple chart types (pie, bar, line)
3. Keyword highlighting
4. AI insights section
5. Icons and status badges
6. Performance meters

Run this to see the enhanced features in action!
"""

import sys
from pathlib import Path

# Add src to path
script_dir = Path(__file__).parent
src_dir = script_dir / 'src'
sys.path.insert(0, str(src_dir))

from pptx import Presentation
from pptx.util import Inches
import pandas as pd
import json

# Import our enhancement modules
from converter.visual_enhancements import VisualEnhancer
from converter.enhanced_charts import EnhancedChartBuilder

# Sample template colors (Corporate Blue)
TEMPLATE_COLORS = {
    'navy': (31, 78, 120),
    'blue': (79, 129, 189),
    'light_blue': (155, 194, 230),
    'white': (255, 255, 255),
    'gray': (127, 127, 127),
    'light_gray': (217, 217, 217),
    'dark_text': (51, 51, 51),
    'chart_colors': [
        (79, 129, 189),
        (194, 57, 52),
        (155, 187, 89),
        (128, 100, 162),
        (75, 172, 198),
        (247, 150, 70)
    ]
}

def main():
    print("\n" + "="*60)
    print("🎨 PROFESSIONAL PPT ENHANCEMENTS DEMO")
    print("="*60)
    
    # Create presentation
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    
    # Initialize enhancers
    visual_enhancer = VisualEnhancer(TEMPLATE_COLORS)
    chart_builder = EnhancedChartBuilder(TEMPLATE_COLORS)
    
    # ===== SLIDE 1: ENHANCED COVER PAGE =====
    print("\n📄 Creating enhanced cover page...")
    visual_enhancer.create_enhanced_cover_slide(
        prs,
        project_name="Q3 2025 Financial Performance",
        subtitle="Executive Summary & Strategic Insights",
        data_period="July - September 2025"
    )
    print("✅ Cover page created with branding!")
    
    # ===== SLIDE 2: PIE CHART DEMO =====
    print("\n📊 Creating pie chart slide...")
    slide2 = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Title
    title_box = slide2.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = "Portfolio Allocation"
    title_frame.paragraphs[0].font.size = Inches(0.4)
    title_frame.paragraphs[0].font.bold = True
    
    # Pie chart data
    allocation_data = {
        'Technology': 35,
        'Healthcare': 25,
        'Finance': 20,
        'Energy': 12,
        'Consumer Goods': 8
    }
    
    chart_builder.create_pie_chart(
        slide2,
        left=1, top=1.5, width=8, height=5,
        data_dict=allocation_data,
        title="Sector Distribution"
    )
    print("✅ Pie chart created!")
    
    # ===== SLIDE 3: BAR CHART DEMO =====
    print("\n📊 Creating bar chart slide...")
    slide3 = prs.slides.add_slide(prs.slide_layouts[6])
    
    title_box = slide3.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = "Top Performers - Q3 2025"
    title_frame.paragraphs[0].font.size = Inches(0.4)
    title_frame.paragraphs[0].font.bold = True
    
    # Bar chart data
    performers = ['Apple Inc.', 'Microsoft', 'Amazon', 'Google', 'Tesla']
    returns = [28.5, 24.3, 19.8, 17.2, 15.9]
    
    chart_builder.create_bar_chart(
        slide3,
        left=1, top=1.5, width=8, height=5,
        categories=performers,
        values=returns,
        title="YTD Returns (%)"
    )
    
    # Add performance icons
    for i, performer in enumerate(performers[:3]):
        visual_enhancer.add_metric_icon(
            slide3,
            left=0.3,
            top=2.0 + (i * 0.9),
            metric_type='growth',
            size=0.25
        )
    
    print("✅ Bar chart with icons created!")
    
    # ===== SLIDE 4: LINE CHART DEMO =====
    print("\n📈 Creating line chart slide...")
    slide4 = prs.slides.add_slide(prs.slide_layouts[6])
    
    title_box = slide4.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = "Revenue Trend Analysis"
    title_frame.paragraphs[0].font.size = Inches(0.4)
    title_frame.paragraphs[0].font.bold = True
    
    # Line chart data
    quarters = ['Q1 2025', 'Q2 2025', 'Q3 2025']
    series_data = {
        'Revenue': [125000, 142000, 168000],
        'Expenses': [85000, 92000, 98000],
        'Profit': [40000, 50000, 70000]
    }
    
    chart_builder.create_line_chart(
        slide4,
        left=1, top=1.5, width=8, height=4.5,
        categories=quarters,
        series_dict=series_data,
        title="Financial Trend"
    )
    
    print("✅ Line chart created!")
    
    # ===== SLIDE 5: AI INSIGHTS WITH KEYWORD HIGHLIGHTING =====
    print("\n🤖 Creating AI insights slide...")
    
    insights = [
        "Strong portfolio growth of 24.5% observed across all sectors",
        "Technology sector shows excellent performance with minimal risk",
        "Revenue increased by 35% compared to previous quarter",
        "Market trends remain stable despite economic concerns",
        "Profit margins improved significantly reaching new high",
        "Investment strategy shows outstanding results with positive outlook"
    ]
    
    analyst_notes = "The portfolio demonstrates strong resilience with excellent diversification. Continue monitoring technology sector for potential growth opportunities while maintaining risk management protocols."
    
    visual_enhancer.create_ai_insights_slide(
        prs,
        insights_list=insights,
        analyst_notes=analyst_notes
    )
    
    print("✅ AI insights slide created!")
    
    # ===== SLIDE 6: STATUS BADGES & PERFORMANCE METERS =====
    print("\n🎯 Creating status indicators slide...")
    slide6 = prs.slides.add_slide(prs.slide_layouts[6])
    
    title_box = slide6.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = "Performance Dashboard"
    title_frame.paragraphs[0].font.size = Inches(0.4)
    title_frame.paragraphs[0].font.bold = True
    
    # Status badges
    visual_enhancer.create_status_badge(slide6, 1, 1.5, "Excellent", "success")
    visual_enhancer.create_status_badge(slide6, 3, 1.5, "On Track", "success")
    visual_enhancer.create_status_badge(slide6, 5, 1.5, "Monitor", "warning")
    visual_enhancer.create_status_badge(slide6, 7, 1.5, "Action Needed", "danger")
    
    # Performance meters
    visual_enhancer.add_performance_meter(slide6, 1, 2.5, 87, label="Portfolio Performance")
    visual_enhancer.add_performance_meter(slide6, 5, 2.5, 92, label="Risk Management")
    visual_enhancer.add_performance_meter(slide6, 1, 4, 73, label="Market Alignment")
    visual_enhancer.add_performance_meter(slide6, 5, 4, 45, label="Sector Diversity")
    
    # Add metric icons
    visual_enhancer.add_metric_icon(slide6, 0.5, 2.5, 'success', size=0.3)
    visual_enhancer.add_metric_icon(slide6, 4.5, 2.5, 'analytics', size=0.3)
    visual_enhancer.add_metric_icon(slide6, 0.5, 4, 'target', size=0.3)
    visual_enhancer.add_metric_icon(slide6, 4.5, 4, 'warning', size=0.3)
    
    print("✅ Status indicators created!")
    
    # ===== SAVE PRESENTATION =====
    output_dir = script_dir / 'test_output'
    output_dir.mkdir(exist_ok=True)
    output_file = output_dir / 'enhanced_demo.pptx'
    
    prs.save(str(output_file))
    
    print("\n" + "="*60)
    print("🎉 SUCCESS!")
    print("="*60)
    print(f"\n📁 Saved to: {output_file}")
    print("\n📊 Features Demonstrated:")
    print("   ✅ Enhanced cover page with branding")
    print("   ✅ Pie chart (sector distribution)")
    print("   ✅ Bar chart (top performers)")
    print("   ✅ Line chart (trend analysis)")
    print("   ✅ AI insights with keyword highlighting")
    print("   ✅ Status badges (success/warning/danger)")
    print("   ✅ Performance meters")
    print("   ✅ Dynamic icons")
    print("\n🎨 Open the PowerPoint to see all enhancements!")
    print("="*60 + "\n")

if __name__ == '__main__':
    main()
