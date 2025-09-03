from pptx import Presentation
from pptx.util import Inches,Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

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

    # Add Headers And Foramte header
    ##Fills the header row (row 0) with DataFrame column names.
    for col_idx, col_name in enumerate(df.columns):
        cell=table.cell(0,col_idx)
        cell.text=str(col_name)

        #Bold,centered gray background
        p=cell.text_frame.paragraphs[0]
        p.font.bold=True
        p.font.size=Pt(12)
        p.alignment=PP_ALIGN.CENTER
        cell.fill.solid()
        cell.fill.fore_color.rgb=RGBColor(200,200,200)

    # Add Data or Fill In the data
    """ 
    Loops through the DataFrame and fills each table cell with its values.
    df.iloc[row_idx, col_idx] → picks a specific cell from the DataFrame.
    """
    for row_idx in range(rows):
        for col_idx in range(cols):
            val=df.iloc[row_idx,col_idx]

            #Formate numbers with commas if numeric
            if isinstance(val,(int,float)):
                text=f"{val:,.0f}"
            else:
                text=str(val)

            cell=table.cell(row_idx+1,col_idx)
            cell.text=text

            ## Align Text
            p=cell.text_frame.paragraphs[0]
            if isinstance(val,(int,float)):
                p.alignment=PP_ALIGN.RIGHT
            else:
                p.alignment=PP_ALIGN.LEFT
            p.font.size=Pt(11)
    ## AUTO FILL columns width
    for col in table.columns:
        col.width=Inches(9.0/cols)

    return prs

def save_ppt(prs, output_path="examples/demo_PPT/output_demo.pptx"):
    """Save the PPTX to file"""
    prs.save(output_path)
