from converter.excel_reader import read_excel
from converter.ppt_writer import create_ppt,add_table_slide,save_ppt
import os
import pandas as pd
def main():
    file_path=os.path.join("Examples","Portfolio Allocation Data.xlsx")

    ## Read the PnL Sheet
    sheets=pd.read_excel(file_path,sheet_name=None)

    ## Create PPT
    prs=create_ppt()

    for sheet_name, df in sheets.items():
        df = df.loc[:, ~df.columns.str.contains("^Unnamed")]  # drop junk cols
        prs = add_table_slide(prs, df, title=sheet_name)

    # Debug: print to confirm shape
    print("✅ DataFrame loaded")
    print("Columns:", df.columns.tolist())
    print("Shape:", df.shape)
    
    

    ## Slide 1:full Portfolio
    prs=add_table_slide(prs,df,title="Full portfoilio Allocation")

    ## Slide 2: Top # assests ONly
    df_top3=df.head(3) ## First 3 rows
    prs=add_table_slide(prs,df_top3,title=" Top 3 Portfoilio Assests")

    ## Save the PPT
    save_ppt(prs, "examples/demo_PPT/output_demo.pptx")
    print(" PPT generated at examples/output_demo.pptx")

if __name__=="__main__":
    main()