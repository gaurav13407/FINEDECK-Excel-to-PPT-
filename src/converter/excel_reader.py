import sys
import re
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

        raw_df=pd.read_excel(path,sheet_name=sheet_page)
        header_row=find_columns(raw_df,required_keyword=None)

        if header_row>=raw_df.shape[0]-1:
            header_value=raw_df.iloc[header_row].fillna("").tolist()
            df=pd.DataFrame(columns=header_value)
            return df
        #promote that row to header
        df=raw_df.iloc[header_row+1:].reset_index(drop=True)
        df.columns=raw_df.iloc[header_row].fillna("").tolist()
        df=clean_columns(df)

        # drop rows to header and get data below it
        df=df.dropna(how="all").reset_index(drop=True)

        # try to coerc numeric columns where it make sense
        for col in df.columns:
            try:
                sample=df[col].astype(str).str.strip().replace({"nan":""})
                non_empty=sample[sample!=""]
                if len(non_empty)==0:
                    continue
                num_like=non_empty.str.match(r'^-?\d+(\.\d+)?$').sum()
                if num_like>=max(1,int(0.4 * len(non_empty))):
                    df[col]=pd.to_numeric(df[col],errors="coerce")
            except Exception as e:
                continue
            
        
        
        return df

    except Exception as e:
        raise CustomException(e,sys)
    

def find_columns(df:pd.DataFrame,required_keyword=None)->int :
    """ Find the first row that contains at least two of the required_keywords.
    If none found, return 0.
    """
    try:
        if required_keyword is None:
            #convert row to lowercase strings,ignore NaN
            logging.info("Findind the Head or where the tables start of the table")
            required_keyword=["asset","value","sector","country","ticker","symbol","amount","price"]
        required_keyword=[k.lower() for k in required_keyword]

        for idx, row in df.iterrows():
            #convert at least 1 or 2 matches depending on sheet shape
            tokens=row.astype(str).str.lower().fillna("")
            matches=sum(any(k in cell for k in required_keyword)for cell in tokens)

            if matches>=1:
                return idx
        return 0
    except Exception as e:
        raise CustomException(e,sys)

def clean_columns(df:pd.DataFrame)->pd.DataFrame:
    #drop fully empty columns
    df=df.dropna(axis=1,how="all")

    #drop columns whoes name starts with Unnamed(after header promotion)
    df=df.iloc[:,[not (isinstance(c,str) and c.strip().lower().startswith("unnamed"))for c in df.columns]]

    #strip whitespace from column names
    new_cols=[]
    for c in df.columns:
        if isinstance(c,str):
            c=c.strip()
            # shorten long name like '2.Portfolio Allocation Data'-> Portfolio Allocation Data
            c=re.sub(r'^\d+\.\s*','',c)
        new_cols.append(c)
    df.columns=new_cols
    return df

if __name__=="__main__":
    file_path = r"examples\finance_sample.xlsx"
    aa=excel_reader(path=file_path,sheet=None)
    print(aa)
