import pandas as pd
import openpyxl
import math
from typing import Dict,Optional,Union
from pptx import Presentation
from pptx.util import Inches,Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from __future__ import annotations

def create_presentation(title:str ="Report",subtitle:str="") -> Presentation:
    prs=Presentation()
    # Title slide layout is usually 0(depends upon template)
    title_slide_layout=prs.slide_layouts[0]
    slide=prs.slides.add_slide(title_slide_layout)
    title_placeholder=slide.shapes.title
    subtitle_placeholder=slide.placeholders[1] if len(slide.placeholders)>1 else None

    title_placeholder.text=title
    if subtitle and subtitle_placeholder:
        subtitle_placeholder.text=subtitle
    return prs

def set_shape_text(shape,text:str,font_size:int=18,bold:bool=False):
    tx=shape.text_frame
    tx.clear()
    p=tx.paragraphs[0]
    run=p.add_run()
    run.text=str(text)
    run.font.size=Pt(font_size)
    run.font.bold=bold

# ---------Slide creators----------------------------------------------------------------------

def row_to_text_slide(prs:Presentation, row:pd.Series, title_col:Optional[str]=None):
    """ 
    Create a text slide for a single row:
    -title:value form title_col (if provided ) or first column
    -boby: bullet list of keys :value for the columns
    """
    layout=prs.slide_layout[1]
    slide=prs.slide.add_slide(layout)

    #Title
    title_text=str(row[title_col]) if title_col and title_col in row.index else str(row.index[0])
    slide.shapes.title.text=title_text

    ## Body placeholder
    body=slide.shapes.placeholders[1].text_frame
    body.clear()

    for col in row.index:
        if title_col and col == title_col:
            continue
        val=row[col]
        if pd.isna(val) or str(val).strip()=="":
            continue
        p=body.add_paragraph()
        p.level=0
        p.text=f"{col}:{val}"

        