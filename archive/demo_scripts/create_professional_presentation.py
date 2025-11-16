"""
Professional Financial Presentation Generator
Creates slides with BOTH data tables AND charts, properly formatted
"""

import sys
import os
sys.path.insert(0, 'src')

from converter.excel_reader import excel_reader, get_sheet_names
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from converter.ppt_writer import (
    create_presentation, 
    create_pie_chart_slide, 
    create_line_chart_slide,
    create_column_chart_slide
)
from converter.chart_detector import detect_chart_type, should_create_chart


def create_data_table_slide(prs, df, title):
    """Create a slide with a data table showing actual values"""
    slide = prs.slides.add_slide(prs.slide_layouts[5])  # Blank layout
    
    # Add title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.5))
    title_box.text = title
    for paragraph in title_box.text_frame.paragraphs:
        paragraph.font.size = Pt(28)
        paragraph.font.bold = True
    
    # Limit rows for readability
    display_df = df.head(10) if len(df) > 10 else df
    
    # Create table
    rows = len(display_df) + 1  # +1 for header
    cols = min(len(display_df.columns), 6)  # Max 6 columns for readability
    
    left = Inches(0.5)
    top = Inches(1.2)
    width = Inches(9)
    height = Inches(5)
    
    table = slide.shapes.add_table(rows, cols, left, top, width, height).table
    
    # Set column widths
    col_width = width / cols
    for i in range(cols):
        table.columns[i].width = col_width
    
    # Header row
    for col_idx in range(cols):
        cell = table.cell(0, col_idx)
        cell.text = str(display_df.columns[col_idx])
        # Format header
        for paragraph in cell.text_frame.paragraphs:
            paragraph.font.bold = True
            paragraph.font.size = Pt(11)
            paragraph.alignment = PP_ALIGN.CENTER
        # Header background color
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(79, 129, 189)  # Blue
        # Text color white
        for paragraph in cell.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.color.rgb = RGBColor(255, 255, 255)
    
    # Data rows
    for row_idx, (_, row) in enumerate(display_df.iterrows(), start=1):
        for col_idx in range(cols):
            cell = table.cell(row_idx, col_idx)
            value = row.iloc[col_idx]
            
            # Format value based on type
            if isinstance(value, (int, float)):
                if abs(value) >= 1_000_000_000:
                    cell.text = f"${value/1_000_000_000:.2f}B"
                elif abs(value) >= 1_000_000:
                    cell.text = f"${value/1_000_000:.2f}M"
                elif abs(value) >= 1000:
                    cell.text = f"${value/1000:.1f}K"
                else:
                    cell.text = f"{value:.2f}"
            else:
                cell.text = str(value)[:30]  # Truncate long text
            
            # Format cell
            for paragraph in cell.text_frame.paragraphs:
                paragraph.font.size = Pt(10)
                paragraph.alignment = PP_ALIGN.CENTER
            
            # Alternate row colors
            if row_idx % 2 == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(242, 242, 242)  # Light gray
    
    return slide


def create_combined_slide(prs, df, title, chart_type, config):
    """Create a slide with BOTH a chart and a data summary table"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank
    
    # Add title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.5))
    title_box.text = title
    for paragraph in title_box.text_frame.paragraphs:
        paragraph.font.size = Pt(24)
        paragraph.font.bold = True
    
    # Left side: Chart (60% width)
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    
    chart_left = Inches(0.5)
    chart_top = Inches(1.0)
    chart_width = Inches(5.5)
    chart_height = Inches(5)
    
    try:
        if chart_type == 'pie':
            chart_df = df[[config['category_col'], config['value_col']]].copy()
            chart_df = chart_df.dropna()
            chart_df = chart_df.sort_values(config['value_col'], ascending=False)
            if len(chart_df) > 8:
                chart_df = chart_df.head(8)
            
            chart_data = CategoryChartData()
            chart_data.categories = chart_df[config['category_col']].tolist()
            chart_data.add_series('Values', chart_df[config['value_col']].tolist())
            
            chart = slide.shapes.add_chart(
                XL_CHART_TYPE.PIE, chart_left, chart_top, chart_width, chart_height, chart_data
            ).chart
            chart.has_legend = True
            chart.plots[0].has_data_labels = True
            data_labels = chart.plots[0].data_labels
            data_labels.number_format = '0%'
            data_labels.font.size = Pt(10)
            
        elif chart_type == 'line':
            chart_data = CategoryChartData()
            x_values = df[config['x_col']].astype(str).tolist()
            # Limit to 20 points for readability
            if len(x_values) > 20:
                step = len(x_values) // 20
                x_values = x_values[::step]
                df_sampled = df.iloc[::step]
            else:
                df_sampled = df
                
            chart_data.categories = x_values
            for y_col in config['y_cols'][:2]:  # Max 2 series for clarity
                if y_col in df_sampled.columns:
                    chart_data.add_series(y_col, df_sampled[y_col].tolist())
            
            chart = slide.shapes.add_chart(
                XL_CHART_TYPE.LINE_MARKERS, chart_left, chart_top, chart_width, chart_height, chart_data
            ).chart
            chart.has_legend = True
            chart.legend.font.size = Pt(10)
            chart.legend.include_in_layout = False
            
        elif chart_type == 'column':
            chart_data = CategoryChartData()
            chart_data.categories = df[config['x_col']].astype(str).tolist()
            for y_col in config['y_cols'][:3]:  # Max 3 series
                if y_col in df.columns:
                    chart_data.add_series(y_col, df[y_col].tolist())
            
            chart = slide.shapes.add_chart(
                XL_CHART_TYPE.COLUMN_CLUSTERED, chart_left, chart_top, chart_width, chart_height, chart_data
            ).chart
            chart.has_legend = True
            chart.legend.font.size = Pt(10)
            chart.legend.include_in_layout = False
    
    except Exception as e:
        print(f"   Chart error: {e}")
    
    # Right side: Data table (35% width)
    table_left = Inches(6.2)
    table_top = Inches(1.0)
    table_width = Inches(3.5)
    
    # Show top 10 rows
    display_df = df.head(10)
    
    # Select key columns to display
    if chart_type == 'pie':
        display_cols = [config['category_col'], config['value_col']]
    elif chart_type == 'line':
        display_cols = [config['x_col']] + config['y_cols'][:2]
    elif chart_type == 'column':
        display_cols = [config['x_col']] + config['y_cols'][:2]
    else:
        display_cols = display_df.columns.tolist()[:3]
    
    display_cols = [col for col in display_cols if col in display_df.columns]
    display_df = display_df[display_cols]
    
    # Create compact table
    rows = min(len(display_df) + 1, 11)  # Max 10 data rows + header
    cols = len(display_cols)
    
    table = slide.shapes.add_table(rows, cols, table_left, table_top, table_width, Inches(5)).table
    
    # Header
    for col_idx, col_name in enumerate(display_cols):
        cell = table.cell(0, col_idx)
        cell.text = str(col_name)[:15]  # Truncate long names
        for paragraph in cell.text_frame.paragraphs:
            paragraph.font.bold = True
            paragraph.font.size = Pt(9)
            paragraph.alignment = PP_ALIGN.CENTER
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(79, 129, 189)
        for paragraph in cell.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.color.rgb = RGBColor(255, 255, 255)
    
    # Data
    for row_idx, (_, row) in enumerate(display_df.iterrows(), start=1):
        if row_idx >= rows:
            break
        for col_idx, col_name in enumerate(display_cols):
            cell = table.cell(row_idx, col_idx)
            value = row[col_name]
            
            # Format value
            if isinstance(value, float):
                if abs(value) >= 1_000_000_000:
                    cell.text = f"{value/1_000_000_000:.1f}B"
                elif abs(value) >= 1_000_000:
                    cell.text = f"{value/1_000_000:.1f}M"
                elif abs(value) >= 1000:
                    cell.text = f"{value/1000:.0f}K"
                else:
                    cell.text = f"{value:.1f}"
            elif isinstance(value, int):
                if abs(value) >= 1_000_000_000:
                    cell.text = f"{value/1_000_000_000:.1f}B"
                elif abs(value) >= 1_000_000:
                    cell.text = f"{value/1_000_000:.1f}M"
                else:
                    cell.text = f"{value:,}"
            else:
                cell.text = str(value)[:12]
            
            for paragraph in cell.text_frame.paragraphs:
                paragraph.font.size = Pt(8)
                paragraph.alignment = PP_ALIGN.RIGHT if isinstance(value, (int, float)) else PP_ALIGN.LEFT
            
            if row_idx % 2 == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(242, 242, 242)
    
    return slide


def create_professional_aapl_presentation():
    """Create a professional presentation with data + charts"""
    
    excel_file = "example/Company_Data/AAPL_Financial_Data.xlsx"
    output_file = "examples/demo_PPT/AAPL_Professional_Complete.pptx"
    
    print("="*70)
    print("📊 PROFESSIONAL AAPL PRESENTATION")
    print("   With Data Tables + Charts + Proper Formatting")
    print("="*70 + "\n")
    
    sheets = get_sheet_names(excel_file)
    prs = create_presentation(
        title="Apple Inc. (AAPL)",
        subtitle="Comprehensive Financial Analysis with Data & Visualizations"
    )
    
    charts_created = 0
    
    for sheet_name in sheets:
        print(f"📄 Processing: {sheet_name}")
        
        try:
            df = excel_reader(excel_file, sheet=sheet_name)
            
            if df is None or df.empty:
                print(f"   ⚠️  Empty\n")
                continue
            
            # Sample large datasets
            if df.shape[0] > 50:
                df = df.tail(30)
            
            if not should_create_chart(df):
                print(f"   ⚠️  Not chartable\n")
                continue
            
            chart_type, config = detect_chart_type(df)
            
            if chart_type is None:
                print(f"   ⚠️  No chart type\n")
                continue
            
            print(f"   ✅ Creating {chart_type.upper()} with data table")
            
            # Create combined slide (chart + data)
            slide = create_combined_slide(
                prs, df, 
                f"AAPL - {sheet_name}",
                chart_type, 
                config
            )
            
            if slide:
                charts_created += 1
                print(f"   🎉 Slide #{charts_created + 1} created!\n")
            
        except Exception as e:
            print(f"   ❌ Error: {str(e)}\n")
    
    # Save
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    prs.save(output_file)
    
    print("="*70)
    print(f"✅ COMPLETE!")
    print("="*70)
    print(f"\n📊 Created: {charts_created + 1} slides (1 title + {charts_created} data+chart slides)")
    print(f"💾 Saved: {output_file}")
    print(f"\n✨ Features:")
    print(f"   • Charts with proper spacing")
    print(f"   • Data tables showing actual values")
    print(f"   • Formatted numbers (B/M/K)")
    print(f"   • Professional styling")
    print(f"   • No overlapping text")
    print("="*70 + "\n")
    
    return output_file


if __name__ == "__main__":
    output = create_professional_aapl_presentation()
    print(f"🎉 Open {output} to see your professional presentation!")
