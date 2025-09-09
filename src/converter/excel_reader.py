import sys
sys.path.append(r"C:\Users\gaura\OneDrive\Desktop\Big Projects\FinDeck(Excel to PPT Project)")
sys.path.append(r"C:\Users\gaura\OneDrive\Desktop\Big Projects\FinDeck(Excel to PPT Project)")
import pandas as pd;
import openpyxl
from src.logs.logger import logging
from src.utils.exceptions.exception import CustomException
""" The Function wriiten Below Is to read the excel and and reutrn it in a Dataframe Formate and for the multiple sheet excel file there is a parameter called Sheet of string type it excpeted the name of the specific sheet in a excel file if 
sheet ir present if any sheet name is not given it will fallback to None which Mean it just get whatever the first sheet is present in that excel file
"""
def excel_reader(path: str, sheet : str | int| None=None)-> pd.DataFrame:
    try:
        logging.info("Reading the excel file")
        if(sheet==None):
            sheet_page=0
        
        else:
            sheet_page=sheet
            
        df=pd.read_excel(path,sheet_name=sheet_page)
        print(df.head(5))
        return df

    except Exception as e:
        raise CustomException(e,sys)
    

if __name__=="__main__":
    file_path = r"examples\Portfolio Allocation Data.xlsx"
    aa=excel_reader(path=file_path,sheet=None)
    print(aa)
