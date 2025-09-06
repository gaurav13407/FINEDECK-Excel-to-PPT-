from converter.excel_reader import read_excel_1
from converter.ppt_writer import create_ppt,add_table_slide,save_ppt,excel_to_ppt
from converter.table_router import add_smart_table_slides
from converter.format_tables import add_dataframe_table_slide
import os
import pandas as pd
def main():
    file_path=os.path.join("Examples","finance_sample.xlsx")

    ## Read the PnL Sheet
    # df=pd.read_excel(file_path,sheet_name=None)

    ## Create PPT
    prs=create_ppt()

   

    # Debug: print to confirm shape
    df=read_excel_1(file_path)
   
    
    

    ## Slide 1:full Portfolio
    prs=add_smart_table_slides(prs,df,title="Full portfoilio Allocation")

    ## Slide 2: Top # assests ONly
    df_top3=df.head(3) ## First 3 rows
    # prs=add_smart_table_slides(prs,df_top3,title=" Top 3 Portfoilio Assests")

    ## Save the PPT
    save_ppt(prs, "examples/demo_PPT/output_demo.pptx")
    print("PPT Genreated Successfuly!")
    print(" PPT generated at examples/output_demo.pptx")

if __name__=="__main__":
    main()