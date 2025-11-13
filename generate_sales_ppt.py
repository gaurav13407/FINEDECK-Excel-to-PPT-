import sys
sys.path.insert(0, r'c:\Users\gaura\OneDrive\Desktop\Big Projects\FinDeck(Excel to PPT Project)\src')
sys.path.insert(0, r'c:\Users\gaura\OneDrive\Desktop\Big Projects\FinDeck(Excel to PPT Project)\src\backend\app')

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import json
from datetime import datetime

# Load the analysis
with open('sales_data_analysis.json', 'r') as f:
    analysis = json.load(f)

# Create presentation
prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(7.5)

def add_title_slide(prs, title, subtitle):
    slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Add title
    left = Inches(1)
    top = Inches(2.5)
    width = Inches(8)
    height = Inches(1)
    title_box = slide.shapes.add_textbox(left, top, width, height)
    title_frame = title_box.text_frame
    title_frame.text = title
    title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    title_frame.paragraphs[0].font.size = Pt(44)
    title_frame.paragraphs[0].font.bold = True
    title_frame.paragraphs[0].font.color.rgb = RGBColor(0, 51, 102)
    
    # Add subtitle
    left = Inches(1)
    top = Inches(3.8)
    subtitle_box = slide.shapes.add_textbox(left, top, width, height)
    subtitle_frame = subtitle_box.text_frame
    subtitle_frame.text = subtitle
    subtitle_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    subtitle_frame.paragraphs[0].font.size = Pt(24)
    subtitle_frame.paragraphs[0].font.color.rgb = RGBColor(100, 100, 100)
    
    return slide

def add_content_slide(prs, title, content_items):
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    
    # Add title
    left = Inches(0.5)
    top = Inches(0.5)
    width = Inches(9)
    height = Inches(0.8)
    title_box = slide.shapes.add_textbox(left, top, width, height)
    title_frame = title_box.text_frame
    title_frame.text = title
    title_frame.paragraphs[0].font.size = Pt(32)
    title_frame.paragraphs[0].font.bold = True
    title_frame.paragraphs[0].font.color.rgb = RGBColor(0, 51, 102)
    
    # Add content
    left = Inches(0.8)
    top = Inches(1.5)
    width = Inches(8.4)
    height = Inches(5)
    content_box = slide.shapes.add_textbox(left, top, width, height)
    text_frame = content_box.text_frame
    text_frame.word_wrap = True
    
    for item in content_items:
        p = text_frame.add_paragraph()
        p.text = item
        p.font.size = Pt(18)
        p.level = 0
        p.space_before = Pt(12)
    
    return slide

# Slide 1: Title
add_title_slide(
    prs,
    "Sales Data Analysis",
    f"Comprehensive Insights from {analysis['executive_summary']['total_records']:,} Records"
)

# Slide 2: Executive Summary
summary = analysis['executive_summary']
content = [
    f"📊 Total Records Analyzed: {summary['total_records']:,}",
    f"📋 Total Columns: {summary['total_columns']}",
    f"",
    "Column Breakdown:",
    f"  • Numeric Columns: {summary['numeric_columns']}",
    f"  • Categorical Columns: {summary['categorical_columns']}",
    f"  • Date Columns: {summary['date_columns']}",
    f"  • Identifier Columns: {summary['identifier_columns']}",
]
add_content_slide(prs, "Executive Summary", content)

# Slide 3: Key Metrics
metrics = analysis['key_metrics']
content = ["Key Financial Metrics:", ""]
for metric in metrics:
    content.append(f"💰 {metric['metric_name']}")
    content.append(f"   Total: ${metric['total']:,.2f}" if 'Sales' in metric['metric_name'] or 'Profit' in metric['metric_name'] else f"   Total: {metric['total']:,.0f}")
    content.append(f"   Average: ${metric['average']:,.2f}" if 'Sales' in metric['metric_name'] or 'Profit' in metric['metric_name'] else f"   Average: {metric['average']:,.2f}")
    content.append("")
add_content_slide(prs, "Key Performance Metrics", content)

# Slide 4: Top Categories - Shipping
ship_mode = analysis['top_categories'][0]
content = [f"Analysis of {ship_mode['category_type']} ({ship_mode['unique_count']} types):", ""]
for val in ship_mode['top_values']:
    content.append(f"📦 {val['category']}: {val['count']:,} orders ({val['percentage']}%)")
add_content_slide(prs, "Shipping Analysis", content)

# Slide 5: Geographic Distribution
if 'hierarchy_analysis' in analysis and 'geographic_hierarchy' in analysis['hierarchy_analysis']:
    geo = analysis['hierarchy_analysis']['geographic_hierarchy']
    content = ["Top Performing Regions:", ""]
    
    if 'top_states' in geo:
        content.append("🌎 Top States by Sales:")
        for state in geo['top_states'][:5]:
            content.append(f"   • {state['location']}: ${state['sales']:,.2f}")
        content.append("")
    
    if 'top_cities' in geo:
        content.append("🏙️ Top Cities by Sales:")
        for city in geo['top_cities'][:5]:
            content.append(f"   • {city['location']}: ${city['sales']:,.2f}")
    
    add_content_slide(prs, "Geographic Performance", content)

# Slide 6: Product Categories
if len(analysis['top_categories']) > 2:
    category = analysis['top_categories'][2]  # Category column
    content = [f"Product {category['category_type']} Distribution:", ""]
    for val in category['top_values'][:7]:
        content.append(f"📦 {val['category']}: {val['count']:,} orders ({val['percentage']}%)")
    add_content_slide(prs, "Product Analysis", content)

# Slide 7: Recommendations
if 'recommendations' in analysis:
    content = ["Strategic Recommendations:", ""]
    for i, rec in enumerate(analysis['recommendations'], 1):
        content.append(f"{i}. {rec}")
    add_content_slide(prs, "Recommendations", content)

# Slide 8: Thank You
add_title_slide(prs, "Thank You", "Data-Driven Insights for Better Decision Making")

# Save presentation
output_file = 'DV_Sales_Data_Analysis.pptx'
prs.save(output_file)
print(f"\n✅ PowerPoint Created Successfully!")
print(f"📄 File: {output_file}")
print(f"📊 Slides: {len(prs.slides)}")
print(f"\nPresentation includes:")
print("  1. Title Slide")
print("  2. Executive Summary")
print("  3. Key Performance Metrics")
print("  4. Shipping Analysis")
print("  5. Geographic Performance")
print("  6. Product Analysis")
print("  7. Strategic Recommendations")
print("  8. Closing Slide")
