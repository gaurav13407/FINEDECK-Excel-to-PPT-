from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import pandas as pd
import traceback

def create_ppt():
    """Creates a blank PPT Presentation"""
    return Presentation()

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import pandas as pd
import traceback

def add_table_slide(prs, df, title="Report", left_margin=Inches(0.5), top_margin=Inches(1.5), table_height=Inches(4)):
    """
    Robust add-table slide:
    - Sanitizes DataFrame
    - If no usable columns, creates a friendly slide describing the problem
    - Falls back to a simple textbox if table creation fails (no raw exception put on slide)
    - Uses actual slide width for sizing
    """
    # Defensive copy
    df = df.copy()

    # Debug info (console only)
    try:
        print(f"DEBUG: Preparing slide '{title}' with shape: {df.shape}")
        if not df.empty:
            print(df.head().to_string())
    except Exception:
        pass

    # If index is meaningful, make it a column (so it appears in PPT)
    try:
        if df.index.name is not None or not all(df.index == range(len(df.index))):
            df = df.reset_index()
    except Exception:
        # if index comparison fails for some reason, just reset index
        df = df.reset_index()

    # Drop "Unnamed" columns only if there are other named columns present
    cols_str = df.columns.astype(str)
    unnamed_mask = cols_str.str.match(r"^Unnamed", na=False)
    if not unnamed_mask.all():
        df = df.loc[:, ~unnamed_mask]

    # Replace NaN with empty strings (pptx doesn't like None)
    df = df.fillna("")

    # Final shape check
    rows, cols = df.shape
    if rows == 0 or cols == 0:
        # Create a friendly slide telling the user why no table was generated
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
        left = left_margin
        top = Inches(0.8)
        width = prs.slide_width - Inches(1.0)
        height = Inches(3.0)
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        tf.word_wrap = True
        tf.text = (
            f"{title}\n\nNo table was created because the DataFrame has shape {df.shape} "
            "— there are no usable columns. Inspect the sheet for missing headers, merged cells, "
            "or wrong header row."
        )
        return prs

    # Ensure columns are strings and non-empty
    df.columns = [str(c) if str(c).strip() != "" else f"col_{i}" for i, c in enumerate(df.columns)]

    # Add slide with title
    slide_layout = prs.slide_layouts[5]  # Title Only
    slide = prs.slides.add_slide(slide_layout)
    try:
        title_placeholder = slide.shapes.title
        title_placeholder.text = title
    except Exception:
        # fallback to a manually placed title
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), prs.slide_width - Inches(1.0), Inches(0.6))
        tf = title_box.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(20)
        p.font.bold = True
        p.alignment = PP_ALIGN.LEFT

    try:
        # Compute available width from the presentation (respect margins)
        total_table_width = prs.slide_width - Inches(1.0)  # left+right margins = 0.5in each
        table = slide.shapes.add_table(rows + 1, cols, left_margin, top_margin, total_table_width, table_height).table

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

        # Fill data rows with safer numeric formatting
        for r in range(rows):
            for c in range(cols):
                val = df.iat[r, c]
                if isinstance(val, (int,)) and not isinstance(val, bool):
                    text = f"{val:,}"
                elif isinstance(val, float) and not isinstance(val, bool):
                    # show up to 2 decimal places, drop trailing zeros
                    text = ("{:.2f}".format(val)).rstrip("0").rstrip(".")
                    # add thousand separators for the integer part
                    if "." in text:
                        int_part, dec_part = text.split(".")
                        int_part = f"{int(int_part):,}"
                        text = int_part + "." + dec_part
                    else:
                        text = f"{int(float(text)):,}"
                else:
                    text = str(val)
                cell = table.cell(r + 1, c)
                cell.text = text
                p = cell.text_frame.paragraphs[0]
                p.font.size = Pt(11)
                try:
                    if isinstance(val, (int, float)) and not isinstance(val, bool):
                        p.alignment = PP_ALIGN.RIGHT
                    else:
                        p.alignment = PP_ALIGN.LEFT
                except Exception:
                    pass

        # Column width — try proportional widths based on max content length
        try:
            max_lens = [max(df[col].astype(str).map(len).max(), len(str(col))) for col in df.columns]
            total_len = sum(max_lens) or len(max_lens)
            for i, ml in enumerate(max_lens):
                # table.columns[i].width expects EMU (Inches() returns EMU)
                table.columns[i].width = int(total_table_width * (ml / total_len))
        except Exception:
            # fallback: equal widths
            try:
                width_per_col = int(total_table_width / cols)
                for col in table.columns:
                    col.width = width_per_col
            except Exception:
                pass

    except Exception as e:
        # Console-only debug; keep slides user-friendly (no raw exception text)
        print(f"add_table_slide: failed to build fancy table for '{title}': {e}")
        traceback.print_exc()

        # Create a readable textbox fallback
        left = left_margin
        top = top_margin
        width = prs.slide_width - Inches(1.0)
        height = table_height
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        tf.word_wrap = True
        tf.text = f"{title}\n\nCould not render a formatted table. A simplified preview follows. Check console logs for details."

        # Provide a short readable preview (not raw exception text)
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

def df_to_table(slide, df, left, top, width, height, include_index=False):
    """Insert a pandas DataFrame into a slide as a pptx table.
    - slide: pptx.slide.Slide
    - df: pandas.DataFrame
    - left, top, width, height: pptx.util.Length values
    Returns the created table shape.
    """
    # Defensive copy
    df = df.copy()

    # Optionally include index as first column
    if include_index:
        df = df.reset_index()

    # Ensure no completely unnamed columns (drop only Unnamed* when mixed)
    cols_str = df.columns.astype(str)
    unnamed_mask = cols_str.str.match(r"^Unnamed", na=False)
    if not unnamed_mask.all():
        df = df.loc[:, ~unnamed_mask]

    # Replace NaN with empty strings
    df = df.fillna("")

    rows, cols = df.shape
    if rows == 0 or cols == 0:
        raise ValueError("Empty dataframe supplied to df_to_table")

    # Force column names to strings
    df.columns = [str(c) if str(c).strip() != "" else f"col_{i}" for i, c in enumerate(df.columns)]

    # create table (header row + data rows)
    table_shape = slide.shapes.add_table(rows + 1, cols, left, top, width, height)
    table = table_shape.table

    # Header
    for c_idx, col_name in enumerate(df.columns):
        cell = table.cell(0, c_idx)
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

    # Fill body
    for r in range(rows):
        for c in range(cols):
            val = df.iat[r, c]
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                # keep decimals if exist; if integer-like, show without decimals
                if float(val).is_integer():
                    text = f"{int(val):,}"
                else:
                    text = f"{val:,}"
            else:
                text = str(val)
            cell = table.cell(r + 1, c)
            cell.text = text
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(11)
            try:
                if isinstance(val, (int, float)) and not isinstance(val, bool):
                    p.alignment = PP_ALIGN.RIGHT
                else:
                    p.alignment = PP_ALIGN.LEFT
            except Exception:
                pass

    # Set column widths proportionally by max string length
    try:
        max_lens = [max(df[col].astype(str).map(len).max(), len(str(col))) for col in df.columns]
        total = sum(max_lens) or len(max_lens)
        for i, ml in enumerate(max_lens):
            table.columns[i].width = int(width * (ml / total))
    except Exception:
        # fallback: equal widths
        try:
            width_per_col = int(width / cols)
            for col in table.columns:
                col.width = width_per_col
        except Exception:
            pass

    return table_shape

def excel_to_ppt(excel_path, ppt_path, rows_per_slide=25, include_index=False):
    """Convert every sheet in excel_path into slides and save to ppt_path.
    Parameters:
    - excel_path: path to .xlsx/.xls file
    - ppt_path: destination .pptx path
    - rows_per_slide: max number of dataframe rows per slide (excluding header)
    - include_index: include DataFrame index as a column
    """
    xls = pd.ExcelFile(excel_path, engine="openpyxl")
    sheet_names = xls.sheet_names

    prs = Presentation()

    left = Inches(0.5)
    top = Inches(1.0)
    slide_width = prs.slide_width - Inches(1.0)
    slide_height = prs.slide_height - Inches(1.6)

    for sheet_name in sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, engine="openpyxl")

        # If empty sheet -> add a simple slide and continue
        if df.shape[0] == 0 or df.shape[1] == 0:
            slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
            add_title(slide, f"{sheet_name} (empty)")
            continue

        # Split into chunks/pages
        chunks = math.ceil(df.shape[0] / rows_per_slide) if rows_per_slide > 0 else 1
        for i in range(chunks):
            start = i * rows_per_slide
            end = min((i + 1) * rows_per_slide, df.shape[0]) if rows_per_slide > 0 else df.shape[0]
            sub_df = df.iloc[start:end]

            slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank slide layout
            page_title = sheet_name if chunks == 1 else f"{sheet_name} — page {i+1}/{chunks}"
            add_title(slide, page_title)

            table_top = top
            table_height = slide_height - Inches(0.6)

            try:
                df_to_table(slide, sub_df, left, table_top, slide_width, table_height, include_index=include_index)
            except Exception as e:
                # fallback: insert a textbox with raw text
                print(f"df_to_table failed for sheet={sheet_name}, page={i+1}: {e}")
                traceback.print_exc()
                txBox = slide.shapes.add_textbox(left, table_top, slide_width, table_height)
                tf = txBox.text_frame
                tf.word_wrap = True
                tf.text = page_title + "\n\n"
                max_rows = min(len(sub_df), 50)
                for r in range(max_rows):
                    row_vals = [str(sub_df.iat[r, c]) for c in range(sub_df.shape[1])]
                    line = " | ".join(row_vals)
                    p = tf.add_paragraph()
                    p.text = line
                    p.font.size = Pt(10)
                if len(sub_df) > max_rows:
                    p = tf.add_paragraph()
                    p.text = "... (truncated)"

    prs.save(ppt_path)
    print(f"PPT saved to: {ppt_path}")


def save_ppt(prs, output_path="examples/demo_PPT/output_demo.pptx"):
    """Save the PPTX to file"""
    prs.save(output_path)
    print(f"PPT saved to: {output_path}")