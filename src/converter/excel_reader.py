import pandas as pd;
import openpyxl
from src.logs.logger import logging
# from utils.exceptions.exception import CustomException 
import sys
""" The Function wriiten Below Is to read the excel and and reutrn it in a Dataframe Formate and for the multiple sheet excel file there is a parameter called Sheet of string type it excpeted the name of the specific sheet in a excel file if 
sheet ir present if any sheet name is not given it will fallback to None which Mean it just get whatever the first sheet is present in that excel file
"""
def excel_reader(path: str, sheet : str | int| None=None)-> pd.DataFrame:
    try:
        logging.info("Reading the excel file")
        df=pd.read_excel(path,sheet_name=sheet)
        print(df.head(5))

    except Exception as e:
        raise (e,sys)
    

if __name__=="__main__":
    excel_reader("examples\Portfolio Allocation Data.xlsx",sheet=None)
    