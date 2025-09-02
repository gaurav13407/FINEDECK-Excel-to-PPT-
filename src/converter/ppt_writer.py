from pptx import Presentation
from pptx.util import Inches

def create_ppt():
    """Creates a blank PPT Presentation"""
    return Presentation()

def add_table_slide(prs, df, title="Report"):
    """Add a slide with a table from a pandas DataFrame"""
    slide_layout = prs.slide_layouts[5]  # Title Only
    slide = prs.slides.add_slide(slide_layout)

    # Add Title
    title_placeholder = slide.shapes.title
    title_placeholder.text = title## Set title at the top of the slide

    # Table dimensions
    """  
    Creates a table shape on the slide.
    rows+1 → adds an extra row for column headers.
    (Inches(0.5), Inches(1.5), Inches(9), Inches(4)) → x, y, width, height of the table.
    """
    rows, cols = df.shape
    table = slide.shapes.add_table(
        rows + 1, cols, Inches(0.5), Inches(1.5), Inches(9), Inches(4)
    ).table

    # Add Headers
    ##Fills the header row (row 0) with DataFrame column names.
    for col_idx, col_name in enumerate(df.columns):
        table.cell(0, col_idx).text = str(col_name)

    # Add Data
    """ 
    Loops through the DataFrame and fills each table cell with its values.
    df.iloc[row_idx, col_idx] → picks a specific cell from the DataFrame.
    """
    for row_idx in range(rows):
        for col_idx in range(cols):
            table.cell(row_idx + 1, col_idx).text = str(df.iloc[row_idx, col_idx])

    return prs

def save_ppt(prs, output_path="examples/demo_PPT/output_demo.pptx"):
    """Save the PPTX to file"""
    prs.save(output_path)
