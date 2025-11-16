"""
Visual Chart Comparison Test
Creates presentations showing the difference between enhanced and basic charts
"""

import sys
import os
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent / 'src'))

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from converter.enhanced_charts import EnhancedChartBuilder
import pandas as pd

def create_comparison_ppt():
    """Create PPT showing enhanced chart features"""
    
    print("\n" + "="*80)
    print("🎨 CREATING VISUAL CHART COMPARISON")
    print("="*80 + "\n")
    
    prs = Presentation()
    prs.slide_width = 9144000
    prs.slide_height = 6858000
    
    # Template colors (Royal Purple)
    template_colors = {
        'navy': (74, 20, 140),
        'blue': (123, 31, 162),
        'light_blue': (206, 147, 216),
        'green': (46, 125, 50),
        'orange': (255, 87, 34),
        'purple': (123, 31, 162),
        'red': (211, 47, 47),
        'yellow': (255, 193, 7),
        'gray': (127, 140, 141),
        'dark_gray': (52, 73, 94),
        'light_gray': (236, 239, 241),
        'white': (255, 255, 255),
        'chart_colors': [(74, 20, 140), (123, 31, 162), (156, 39, 176), (186, 104, 200), (206, 147, 216), (225, 190, 231)]
    }
    
    chart_builder = EnhancedChartBuilder(template_colors)
    
    # =========================================================================
    # SLIDE 1: Title
    # =========================================================================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Background
    bg = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(*template_colors['navy'])
    bg.line.color.rgb = RGBColor(*template_colors['navy'])
    
    # Title
    title_box = slide.shapes.add_textbox(Inches(1), Inches(2.5), Inches(8), Inches(2))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = "Enhanced Chart\nVisualization Demo"
    p.font.size = Pt(54)
    p.font.bold = True
    p.font.color.rgb = RGBColor(*template_colors['white'])
    p.alignment = 1  # Center
    
    subtitle_box = slide.shapes.add_textbox(Inches(1), Inches(4.8), Inches(8), Inches(0.8))
    tf = subtitle_box.text_frame
    p = tf.paragraphs[0]
    p.text = "See the Difference: Auto-Detection & Template Colors"
    p.font.size = Pt(24)
    p.font.color.rgb = RGBColor(*template_colors['light_blue'])
    p.alignment = 1  # Center
    
    print("✅ Created title slide")
    
    # =========================================================================
    # SLIDE 2: PIE CHART - Portfolio Allocation
    # =========================================================================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = "📊 Enhanced Pie Chart - Portfolio Allocation"
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = RGBColor(*template_colors['navy'])
    
    # Sample data
    allocation_data = {
        'Stocks': 45,
        'Bonds': 25,
        'Real Estate': 15,
        'Cash': 10,
        'Commodities': 5
    }
    
    chart_builder.create_pie_chart(
        slide,
        left=1,
        top=1.5,
        width=8,
        height=5,
        data_dict=allocation_data,
        title="Portfolio Allocation"
    )
    
    # Add description
    desc_box = slide.shapes.add_textbox(Inches(0.5), Inches(6.7), Inches(9), Inches(0.5))
    tf = desc_box.text_frame
    p = tf.paragraphs[0]
    p.text = "✨ Features: Purple template colors • Percentage labels • Professional legend • Auto-detected as PIE for distribution data"
    p.font.size = Pt(14)
    p.font.color.rgb = RGBColor(*template_colors['dark_gray'])
    p.font.italic = True
    
    print("✅ Created pie chart slide")
    
    # =========================================================================
    # SLIDE 3: BAR CHART - Top Performers
    # =========================================================================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = "📊 Enhanced Bar Chart - Top Performers"
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = RGBColor(*template_colors['navy'])
    
    # Sample data
    categories = ['Product A', 'Product B', 'Product C', 'Product D', 'Product E']
    values = [95, 87, 82, 76, 68]
    
    chart_builder.create_bar_chart(
        slide,
        left=1,
        top=1.5,
        width=8,
        height=5,
        categories=categories,
        values=values,
        title="Top 5 Products by Revenue"
    )
    
    # Add description
    desc_box = slide.shapes.add_textbox(Inches(0.5), Inches(6.7), Inches(9), Inches(0.5))
    tf = desc_box.text_frame
    p = tf.paragraphs[0]
    p.text = "✨ Features: Horizontal bars • Purple gradient • Value labels • Auto-detected as BAR for ≤10 categories"
    p.font.size = Pt(14)
    p.font.color.rgb = RGBColor(*template_colors['dark_gray'])
    p.font.italic = True
    
    print("✅ Created bar chart slide")
    
    # =========================================================================
    # SLIDE 4: LINE CHART - Revenue Trends
    # =========================================================================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = "📊 Enhanced Line Chart - Revenue Trends"
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = RGBColor(*template_colors['navy'])
    
    # Sample data
    periods = ['Q1 2024', 'Q2 2024', 'Q3 2024', 'Q4 2024', 'Q1 2025']
    series = {
        'Revenue': [120, 145, 160, 175, 195],
        'Profit': [25, 32, 38, 45, 52]
    }
    
    chart_builder.create_line_chart(
        slide,
        left=1,
        top=1.5,
        width=8,
        height=5,
        categories=periods,
        series_dict=series,
        title="Quarterly Performance Trends"
    )
    
    # Add description
    desc_box = slide.shapes.add_textbox(Inches(0.5), Inches(6.7), Inches(9), Inches(0.5))
    tf = desc_box.text_frame
    p = tf.paragraphs[0]
    p.text = "✨ Features: Purple lines with markers • Multiple series • Legend • Auto-detected as LINE for time-series data"
    p.font.size = Pt(14)
    p.font.color.rgb = RGBColor(*template_colors['dark_gray'])
    p.font.italic = True
    
    print("✅ Created line chart slide")
    
    # =========================================================================
    # SLIDE 5: COLUMN CHART - Regional Comparison
    # =========================================================================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = "📊 Enhanced Column Chart - Regional Sales"
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = RGBColor(*template_colors['navy'])
    
    # Sample data
    regions = ['North', 'South', 'East', 'West', 'Central']
    regional_series = {
        '2024': [150, 135, 145, 128, 142],
        '2025': [165, 148, 158, 140, 155]
    }
    
    chart_builder.create_column_chart(
        slide,
        left=1,
        top=1.5,
        width=8,
        height=5,
        categories=regions,
        series_dict=regional_series,
        title="Regional Sales Comparison"
    )
    
    # Add description
    desc_box = slide.shapes.add_textbox(Inches(0.5), Inches(6.7), Inches(9), Inches(0.5))
    tf = desc_box.text_frame
    p = tf.paragraphs[0]
    p.text = "✨ Features: Vertical columns • Purple color scheme • Grouped series • Auto-detected as COLUMN for category data"
    p.font.size = Pt(14)
    p.font.color.rgb = RGBColor(*template_colors['dark_gray'])
    p.font.italic = True
    
    print("✅ Created column chart slide")
    
    # =========================================================================
    # SLIDE 6: Summary
    # =========================================================================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Background
    bg = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(*template_colors['light_gray'])
    bg.line.color.rgb = RGBColor(*template_colors['light_gray'])
    
    title_box = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(0.8))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = "✅ Enhanced Chart Features Summary"
    p.font.size = Pt(40)
    p.font.bold = True
    p.font.color.rgb = RGBColor(*template_colors['navy'])
    p.alignment = 1
    
    features = [
        "✨ Auto-Detection: Automatically selects PIE/BAR/LINE/COLUMN based on data type",
        "🎨 Template Colors: All charts use Royal Purple color scheme for brand consistency",
        "📊 Professional Formatting: Better labels, legends, and data presentation",
        "🔄 Intelligent Fallback: Gracefully handles errors with simple chart alternatives",
        "📈 Multiple Series Support: Line and column charts support multiple data series",
        "💯 Percentage Labels: Pie charts show % values for easy interpretation"
    ]
    
    y_pos = 1.8
    for feature in features:
        text_box = slide.shapes.add_textbox(Inches(1), Inches(y_pos), Inches(8), Inches(0.6))
        tf = text_box.text_frame
        p = tf.paragraphs[0]
        p.text = feature
        p.font.size = Pt(18)
        p.font.color.rgb = RGBColor(*template_colors['dark_gray'])
        p.space_after = Pt(12)
        y_pos += 0.8
    
    print("✅ Created summary slide")
    
    # Save
    output_path = "test_output/chart_comparison_demo.pptx"
    os.makedirs("test_output", exist_ok=True)
    prs.save(output_path)
    
    print("\n" + "="*80)
    print("✅ CHART COMPARISON PPT CREATED!")
    print("="*80)
    print(f"📂 Output: {output_path}")
    print("\n📋 What to look for:")
    print("  • Slide 2: PIE chart in PURPLE with percentages")
    print("  • Slide 3: HORIZONTAL BAR chart in PURPLE")
    print("  • Slide 4: LINE chart with PURPLE lines and markers")
    print("  • Slide 5: COLUMN chart with PURPLE vertical bars")
    print("\n🎯 All charts use Royal Purple template colors!")
    print("="*80 + "\n")

if __name__ == "__main__":
    create_comparison_ppt()
