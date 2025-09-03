from converter.excel_reader import read_excel_1
from converter.ppt_writer import create_ppt,add_table_slide,save_ppt
import os
import pandas as pd
def main():
    file_path=os.path.join("Examples","Portfolio Allocation Data.xlsx")

    ## Read the PnL Sheet
    # df=pd.read_excel(file_path,sheet_name=None)

    ## Create PPT
    prs=create_ppt()

   

    # Debug: print to confirm shape
    df=read_excel_1(file_path)
   
    
    

    ## Slide 1:full Portfolio
    prs=add_table_slide(prs,df,title="Full portfoilio Allocation")

    ## Slide 2: Top # assests ONly
    df_top3=df.head(3) ## First 3 rows
    prs=add_table_slide(prs,df_top3,title=" Top 3 Portfoilio Assests")

    ## Save the PPT
    save_ppt(prs, "examples/demo_PPT/output_demo.pptx")
    print("PPT Genreated Successfuly!")
    print(" PPT generated at examples/output_demo.pptx")

if __name__=="__main__":
    main()