# ppt_writer.py
import os
import sys
from typing import Optional, Dict, Any
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# Add parent directory to path for imports
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

from src.templates.template_manager import TemplateManager, hex_to_rgb, get_default_template

# Global template manager instance
_template_manager = None

def get_template_manager():
    """Get or create template manager instance"""
    global _template_manager
    if _template_manager is None:
        _template_manager = TemplateManager()
    return _template_manager

def create_presentation(title: str = "Report", subtitle: str = "", template_name: Optional[str] = None) -> Presentation:
    """
    Create a new presentation with optional template styling
    
    Args:
        title: Presentation title
        subtitle: Presentation subtitle
        template_name: Template to use (default: corporate_blue)
    
    Returns:
        Presentation object with styled title slide
    """
    prs = Presentation()
    
    # Load template if specified
    template = None
    if template_name:
        template = get_template_manager().load_template(template_name)
    
    # Title slide layout is usually 0 (depends on template)
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title_placeholder = slide.shapes.title
    # placeholder index 1 is often subtitle, but check existence
    subtitle_placeholder = slide.placeholders[1] if len(slide.placeholders) > 1 else None

    title_placeholder.text = title
    if subtitle and subtitle_placeholder:
        subtitle_placeholder.text = subtitle
    
    # Apply template styling to title slide
    if template:
        _apply_template_to_title_slide(slide, template)
    
    return prs

def _apply_template_to_title_slide(slide, template: Dict[str, Any]):
    """Apply template styling to title slide"""
    try:
        title_font = template['fonts']['title']
        subtitle_font = template['fonts']['subtitle']
        
        # Style title
        if slide.shapes.title:
            for paragraph in slide.shapes.title.text_frame.paragraphs:
                paragraph.font.size = Pt(title_font['size'])
                paragraph.font.bold = title_font.get('bold', True)
                if 'color' in title_font:
                    r, g, b = hex_to_rgb(title_font['color'])
                    paragraph.font.color.rgb = RGBColor(r, g, b)
        
        # Style subtitle
        if len(slide.placeholders) > 1:
            for paragraph in slide.placeholders[1].text_frame.paragraphs:
                paragraph.font.size = Pt(subtitle_font['size'])
                paragraph.font.bold = subtitle_font.get('bold', False)
                if 'color' in subtitle_font:
                    r, g, b = hex_to_rgb(subtitle_font['color'])
                    paragraph.font.color.rgb = RGBColor(r, g, b)
    except Exception as e:
        print(f"Warning: Could not apply template styling to title: {e}")

def set_shape_text(shape, text: str, font_size: int = 18, bold: bool = False):
    tx = shape.text_frame
    tx.clear()
    p = tx.paragraphs[0]
    run = p.add_run()
    run.text = str(text)
    run.font.size = Pt(font_size)
    run.font.bold = bold

# --------- Slide creators ---------------------------------------------------

def row_to_text_slide(prs: Presentation, row: pd.Series, title_col: Optional[str] = None):
    """
    Create a text slide for a single row:
    - title: value from title_col (if provided) or the first non-empty field
    - body: bullet list of key: value for the other columns
    """
    layout = prs.slide_layouts[1]  # Title + Content layout (common)
    slide = prs.slides.add_slide(layout)

    # determine title text
    if title_col and title_col in row.index and not pd.isna(row[title_col]) and str(row[title_col]).strip() != "":
        title_text = str(row[title_col])
    else:
        # fallback: first non-empty column value or column name
        title_text = None
        for c in row.index:
            if not pd.isna(row[c]) and str(row[c]).strip() != "":
                title_text = str(row[c])
                break
        if title_text is None:
            # ultimate fallback: first column name
            title_text = str(row.index[0])

    # set slide title
    if slide.shapes.title:
        slide.shapes.title.text = title_text

    # get body placeholder (usually index 1)
    body_placeholder = None
    if len(slide.placeholders) > 1:
        body_placeholder = slide.placeholders[1]
    else:
        # try to find a content placeholder among shapes
        for shp in slide.shapes:
            if hasattr(shp, "text_frame"):
                body_placeholder = shp
                break

    if body_placeholder is None:
        # add textbox if no existing body placeholder
        body_placeholder = slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(8), Inches(4))

    body = body_placeholder.text_frame
    body.clear()

    # add bullet lines
    for col in row.index:
        if title_col and col == title_col:
            continue
        val = row[col]
        if pd.isna(val) or str(val).strip() == "":
            continue
        p = body.add_paragraph()
        p.level = 0
        p.text = f"{col}: {val}"

def row_to_table_slide(prs: Presentation, row: pd.Series, title_col: Optional[str] = None, max_cols: int = 2):
    """
    Create a slide with a simple 2-column table: Field | Value
    """
    layout = prs.slide_layouts[5] if len(prs.slide_layouts) > 5 else prs.slide_layouts[1]
    slide = prs.slides.add_slide(layout)

    # Title (if present)
    if slide.shapes.title:
        if title_col and title_col in row.index and not pd.isna(row[title_col]) and str(row[title_col]).strip() != "":
            slide.shapes.title.text = str(row[title_col])
        else:
            for c in row.index:
                if not pd.isna(row[c]) and str(row[c]).strip() != "":
                    slide.shapes.title.text = str(row[c])
                    break

    # Prepare table data
    items = [(col, row[col]) for col in row.index if not pd.isna(row[col]) and str(row[col]).strip() != ""]
    if not items:
        # empty row; add a small textbox
        tx_box = slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(8), Inches(1))
        set_shape_text(tx_box, "No data", font_size=20)
        return

    rows = len(items) + 1  # header + items
    cols = max_cols
    left = Inches(0.5)
    top = Inches(1.8)
    width = Inches(9)
    height = Inches(0.6 + 0.2 * rows)

    table = slide.shapes.add_table(rows, cols, left, top, width, height).table

    # header row
    table.cell(0, 0).text = "Field"
    table.cell(0, 1).text = "Value"

    # fill table
    for i, (k, v) in enumerate(items, start=1):
        table.cell(i, 0).text = str(k)
        table.cell(i, 1).text = str(v)

    # Set font sizes for table cells
    for r in range(rows):
        for c in range(cols):
            for paragraph in table.cell(r, c).text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(12)


#-------------------------- Chart Creator---------------------------------------------------
def create_pie_chart_slide(prs: Presentation, df:pd.DataFrame,category_col:str,value_col:str,title:str="Distribution Chart"):
    """ Create a pie chart slide from DataFrame columns """
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.util import Inches

    # Add Slide with balnk layout 
    slide=prs.slides.add_slide(prs.slide_layouts[6]) # blank layout

    ## Add title
    title_box=slide.shapes.add_textbox(Inches(0.5),Inches(0.2),Inches(9),Inches(0.5))
    title_box.text=title
    for paragraph in title_box.text_frame.paragraphs:
        paragraph.font.size=Pt(24)
        paragraph.font.bold=True

    ## Prepare data-sort by value descending
    chart_df=df[[category_col,value_col]].copy()
    chart_df=chart_df.dropna()
    chart_df=chart_df.sort_values(by=value_col,ascending=False)

    ## Limit to top 10 categories
    if len(chart_df)>10:
        chart_df=chart_df.head(10)

    # Create chart data
    chart_data=CategoryChartData()
    chart_data.categories=chart_df[category_col].tolist()
    chart_data.add_series('Values', chart_df[value_col].tolist())

    x,y,cx,cy=Inches(1),Inches(1.2),Inches(8),Inches(5)
    chart=slide.shapes.add_chart(XL_CHART_TYPE.PIE, x,y,cx,cy,chart_data).chart

    # Chart Styling
    chart.has_legend=True
    from pptx.enum.chart import XL_LEGEND_POSITION
    chart.legend.position=XL_LEGEND_POSITION.RIGHT
    chart.legend.font.size=Pt(10)

    # Show Percentages on pie slices
    chart.plots[0].has_data_labels=True
    data_labels=chart.plots[0].data_labels
    data_labels.number_format='0%'
    data_labels.position=5

    return slide

def create_bar_chart_slide(prs: Presentation, df: pd.DataFrame, x_col: str, 
                           y_col: str, title: str = "Comparison"):
    """
    Create a horizontal bar chart slide.
    Best for: Comparing categories, ranking, metrics comparison
    """
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.util import Inches
    
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Add title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.5))
    title_box.text = title
    for paragraph in title_box.text_frame.paragraphs:
        paragraph.font.size = Pt(24)
        paragraph.font.bold = True
    
    # Prepare data
    chart_df = df[[x_col, y_col]].copy()
    chart_df = chart_df.dropna()
    chart_df = chart_df.sort_values(y_col, ascending=True)  # Ascending for bar chart
    
    # Create chart data
    chart_data = CategoryChartData()
    chart_data.categories = chart_df[x_col].tolist()
    chart_data.add_series(y_col, chart_df[y_col].tolist())
    
    # Add chart
    x, y, cx, cy = Inches(1), Inches(1.2), Inches(8), Inches(5)
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED, x, y, cx, cy, chart_data
    ).chart
    
    # Styling
    chart.has_legend = False
    chart.plots[0].has_data_labels = True
    data_labels = chart.plots[0].data_labels
    data_labels.number_format = '#,##0'
    
    return slide


def create_line_chart_slide(prs: Presentation, df: pd.DataFrame, x_col: str, 
                            y_cols: list, title: str = "Trend Analysis"):
    """
    Create a line chart slide showing trends over time.
    Best for: Time series, quarterly data, trend analysis
    """
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.util import Inches
    
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Add title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.5))
    title_box.text = title
    for paragraph in title_box.text_frame.paragraphs:
        paragraph.font.size = Pt(24)
        paragraph.font.bold = True
    
    # Prepare data
    chart_data = CategoryChartData()
    chart_data.categories = df[x_col].tolist()
    
    # Add series for each numeric column
    for y_col in y_cols[:3]:  # Max 3 series for clarity
        if y_col in df.columns:
            chart_data.add_series(y_col, df[y_col].tolist())
    
    # Add chart
    x, y, cx, cy = Inches(1), Inches(1.2), Inches(8), Inches(5)
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE_MARKERS, x, y, cx, cy, chart_data
    ).chart
    
    # Styling
    chart.has_legend = True
    from pptx.enum.chart import XL_LEGEND_POSITION
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.font.size = Pt(10)
    
    # Add data labels only if single series
    if len(y_cols) == 1:
        chart.plots[0].has_data_labels = True
    
    return slide


def create_column_chart_slide(prs: Presentation, df: pd.DataFrame, x_col: str, 
                              y_cols: list, title: str = "Comparison"):
    """
    Create a column (vertical bar) chart slide.
    Best for: Multi-series comparison, grouped data
    """
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.util import Inches
    
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Add title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.5))
    title_box.text = title
    for paragraph in title_box.text_frame.paragraphs:
        paragraph.font.size = Pt(24)
        paragraph.font.bold = True
    
    # Prepare data
    chart_data = CategoryChartData()
    chart_data.categories = df[x_col].tolist()
    
    # Add series
    for y_col in y_cols[:3]:  # Max 3 series
        if y_col in df.columns:
            chart_data.add_series(y_col, df[y_col].tolist())
    
    # Add chart
    x, y, cx, cy = Inches(1), Inches(1.2), Inches(8), Inches(5)
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data
    ).chart
    
    # Styling
    chart.has_legend = True
    from pptx.enum.chart import XL_LEGEND_POSITION
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.font.size = Pt(10)
    
    return slide


def create_scatter_chart_slide(prs: Presentation, df: pd.DataFrame, x_col: str,
                              y_col: str, title: str = "Correlation Analysis"):
    """
    Create a scatter plot slide.
    Best for: Correlation analysis, relationship between two variables
    """
    from pptx.chart.data import XyChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.util import Inches
    
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Add title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.5))
    title_box.text = title
    for paragraph in title_box.text_frame.paragraphs:
        paragraph.font.size = Pt(24)
        paragraph.font.bold = True
    
    # Prepare data
    chart_df = df[[x_col, y_col]].copy()
    chart_df = chart_df.dropna()
    
    # Create XY chart data
    chart_data = XyChartData()
    series = chart_data.add_series(f'{y_col} vs {x_col}')
    for _, row in chart_df.iterrows():
        series.add_data_point(row[x_col], row[y_col])
    
    # Add chart
    x, y, cx, cy = Inches(1), Inches(1.2), Inches(8), Inches(5)
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.XY_SCATTER, x, y, cx, cy, chart_data
    ).chart
    
    # Styling
    chart.has_legend = False
    
    return slide


def create_auto_chart_slide(prs: Presentation, df: pd.DataFrame, title: str = "Data Visualization"):
    """
    Automatically detect and create the most appropriate chart for the data.
    
    Uses intelligent detection based on:
    - Column types (numeric vs categorical)
    - Data patterns (time series, part-to-whole, comparisons)
    - Row count and data distribution
    """
    from converter.chart_detector import detect_chart_type, should_create_chart
    
    # Check if data is suitable for charting
    if not should_create_chart(df):
        return None
    
    # Detect appropriate chart type
    chart_type, config = detect_chart_type(df)
    
    if chart_type is None:
        return None
    
    # Create the appropriate chart
    slide = None
    try:
        if chart_type == 'pie':
            slide = create_pie_chart_slide(
                prs, df, 
                config['category_col'], 
                config['value_col'],
                config.get('title', title)
            )
        elif chart_type == 'bar':
            slide = create_bar_chart_slide(
                prs, df,
                config['x_col'],
                config['y_col'],
                config.get('title', title)
            )
        elif chart_type == 'line':
            slide = create_line_chart_slide(
                prs, df,
                config['x_col'],
                config['y_cols'],
                config.get('title', title)
            )
        elif chart_type in ['column', 'stacked_bar']:
            slide = create_column_chart_slide(
                prs, df,
                config['x_col'],
                config['y_cols'],
                config.get('title', title)
            )
        elif chart_type == 'scatter':
            slide = create_scatter_chart_slide(
                prs, df,
                config['x_col'],
                config['y_col'],
                config.get('title', title)
            )
    except Exception as e:
        print(f"Warning: Could not create {chart_type} chart: {e}")
        return None
    
    return slide






# --- Main export function --------------------------------------------------

def df_to_ppt(df: pd.DataFrame, out_path: str, title: str = "Auto Report", subtitle: str = "",
              title_col: Optional[str] = None, mode: str = "table", limit: Optional[int] = None,include_charts: bool = True):
    """
    Convert a DataFrame to PPT:
    - mode: 'table' or 'text' (per-row slide)
    - title_col: optional, column name used for slide title
    - limit: optional max number of rows to export (useful for testing)
    """
    if df is None or df.shape[0] == 0:
        prs = create_presentation(title, subtitle)
        layout = prs.slide_layouts[1] if len(prs.slide_layouts) > 1 else prs.slide_layouts[0]
        slide = prs.slides.add_slide(layout)
        if slide.shapes.title:
            slide.shapes.title.text = "No data"
        prs.save(out_path)
        return out_path

    prs = create_presentation(title, subtitle)
    n = df.shape[0]
    if limit is not None:
        n = min(n, int(limit))
    
    # MODE 1: Chart Only - Just create visualization
    if mode == "chart_only":
        create_auto_chart_slide(prs, df.iloc[:n], title="Data Overview")
    
    # MODE 2: Auto - Smart detection
    elif mode == "auto":
        # First, add a chart overview if data is suitable
        if include_charts:
            from converter.chart_detector import should_create_chart
            if should_create_chart(df.iloc[:n]):
                create_auto_chart_slide(prs, df.iloc[:n], title="Overview")
        
        # Then add detailed table/text slides if needed (for small datasets)
        if n <= 10:
            for i in range(n):
                row = df.iloc[i]
                row_to_table_slide(prs, row, title_col=title_col)
    
    # MODE 3: Table mode
    elif mode == "table":
        # Add chart first if suitable
        if include_charts:
            create_auto_chart_slide(prs, df.iloc[:n], title="Overview")
        
        # Add table slides
        for i in range(n):
            row = df.iloc[i]
            row_to_table_slide(prs, row, title_col=title_col)
    
    # MODE 4: Text mode
    elif mode == "text":
        # Add chart first if suitable
        if include_charts:
            create_auto_chart_slide(prs, df.iloc[:n], title="Overview")
        
        # Add text slides
        for i in range(n):
            row = df.iloc[i]
            row_to_text_slide(prs, row, title_col=title_col)

    # Ensure output directory exists
    out_dir = os.path.dirname(out_path)
    if out_dir and not os.path.exists(out_dir):
        os.makedirs(out_dir, exist_ok=True)

    prs.save(out_path)
    return out_path

# --- Example usage ---------------------------------------------------------

if __name__ == "__main__":
    demo_path = os.path.join("example", "AAPL_Financial_Data.xlsx")
    df = pd.read_excel(demo_path, sheet_name=0)
    df = df.fillna("").astype(object)
    out = df_to_ppt(df, out_path=os.path.join("examples", "demo_presentation.pptx"),
                    title="Finance Sample", subtitle="Auto-generated", title_col="Asset", mode="table", limit=10)
    print("Saved:", out)
