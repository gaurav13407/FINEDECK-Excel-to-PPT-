from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import pandas as pd
import traceback

def create_ppt():
    """Creates a blank PPT Presentation"""
    return Presentation()

def add_table_slide(prs, df, title="Report"):
    """
    Robust add-table slide:
    - Sanitizes DataFrame
    - Skips truly empty frames
    - Falls back to a simple textbox if table creation fails
    """
    # Defensive copy
    df = df.copy()

    # Debug info so you can see what's being passed
    try:
        print(f"DEBUG: Preparing slide '{title}' with shape: {df.shape}")
        if not df.empty:
            print(df.head().to_string())
    except Exception:
        pass

    # If index is meaningful, make it a column (so it appears in PPT)
    if df.index.name is not None or not all(df.index == range(len(df.index))):
        df = df.reset_index()

    # Drop "Unnamed" columns only if there are other named columns present
    cols_str = df.columns.astype(str)
    unnamed_mask = cols_str.str.match(r"^Unnamed", na=False)
    if unnamed_mask.all():
        # all columns unnamed -> keep as-is (maybe single column without header)
        pass
    else:
        # drop only unnamed columns if some named columns exist
        df = df.loc[:, ~cols_str.str.match(r"^Unnamed", na=False)]

    # Replace NaN with empty strings (pptx doesn't like None)
    df = df.fillna("")

    # Final shape check
    rows, cols = df.shape
    if rows == 0 or cols == 0:
        print(f"⚠️ Skipping empty DataFrame for slide: {title} (shape={df.shape})")
        return prs

    # Ensure columns are strings and non-empty
    df.columns = [str(c) if str(c).strip() != "" else f"col_{i}" for i, c in enumerate(df.columns)]

    # Add slide with title
    slide_layout = prs.slide_layouts[5]  # Title Only
    slide = prs.slides.add_slide(slide_layout)
    title_placeholder = slide.shapes.title
    title_placeholder.text = title

    try:
        # Create table (rows + header)
        table = slide.shapes.add_table(rows + 1, cols, Inches(0.5), Inches(1.5), Inches(9), Inches(4)).table

        # Header formatting
        for col_idx, col_name in enumerate(df.columns):
            cell = table.cell(0, col_idx)
            cell.text = str(col_name)
            p = cell.text_frame.paragraphs[0]
            p.font.bold = True
            p.font.size = Pt(12)
            p.alignment = PP_ALIGN.CENTER
            try:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(200, 200, 200)
            except Exception:
                pass

        # Fill data rows
        for r in range(rows):
            for c in range(cols):
                val = df.iat[r, c]
                # numeric pretty-format
                if isinstance(val, (int, float)) and not isinstance(val, bool):
                    text = f"{val:,.0f}"
                else:
                    text = str(val)
                cell = table.cell(r + 1, c)
                cell.text = text
                p = cell.text_frame.paragraphs[0]
                try:
                    if isinstance(val, (int, float)) and not isinstance(val, bool):
                        p.alignment = PP_ALIGN.RIGHT
                    else:
                        p.alignment = PP_ALIGN.LEFT
                except Exception:
                    pass
                p.font.size = Pt(11)

        # Safe column width calculation
        if cols > 0:
            try:
                width_per_col = Inches(9.0 / cols)
                for col in table.columns:
                    col.width = width_per_col
            except Exception:
                # ignore width errors
                pass

    except Exception as e:
        # Fallback: if table creation fails, add readable text box
        print("Fancy table failed:", str(e))
        traceback.print_exc()
        left, top, width, height = Inches(0.5), Inches(1.5), Inches(9), Inches(4)
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        tf.word_wrap = True
        try:
            tf.text = f"{title}\n\n"
        except Exception:
            tf.text = title + "\n\n"
        max_rows = min(rows, 20)
        for r in range(max_rows):
            row_vals = [str(df.iat[r, c]) for c in range(cols)]
            line = " | ".join(row_vals)
            p = tf.add_paragraph()
            p.text = line
            p.font.size = Pt(10)
        if rows > max_rows:
            p = tf.add_paragraph()
            p.text = "... (truncated)"

    return prs

def save_ppt(prs, output_path="examples/demo_PPT/output_demo.pptx"):
    """Save the PPTX to file"""
    prs.save(output_path)
    print(f"PPT saved to: {output_path}")
