from typing import Any
from pptx import Presentation

def add_smart_table_slides(prs:Presentation,df,title:str="Report")->Any:
    """ 
    Try advanced styled table (formate_tables functions)
    If Anything fails ,fall back to ppt_writer function.
    """
    try:
        from src.converter.format_tables import add_dataframe_table_slide
        return add_dataframe_table_slide(prs,df,title=title)
    except Exception as e:
        print(f"{e} Fancy table failed:{e}\n Falling back to basic table.")
        from converter.ppt_writer import add_table_slide
        return add_table_slide(prs,df,title=title)