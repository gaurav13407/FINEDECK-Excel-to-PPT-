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

        has_unnamed = any(
        isinstance(c, str) and c.strip().lower().startswith("unnamed")
        for c in raw_df.columns
            )

        if not has_unnamed:
        # No Unnamed headers => skip promotion entirely.
            logging.info("No 'Unnamed' columns detected — skipping header promotion.")
            # Light sanitize columns (trim and remove numeric prefixes), but do NOT promote/alter header rows.
            df = clean_columns(raw_df)   # make sure your clean_columns does light sanitization in this case
            df = df.dropna(how="all").reset_index(drop=True)
            return df
        # print(raw_df)
        print("\n")
        header_row=find_columns(raw_df,required_keyword=None)
        logging.info("Loading the raw data from excel")
        if header_row>=raw_df.shape[0]-1:
            header_value=raw_df.iloc[header_row].fillna("").tolist()
            df=pd.DataFrame(columns=header_value)
            return df
        #promote that row to header
        df=raw_df.iloc[header_row+1:].reset_index(drop=True)
        df.columns=raw_df.iloc[header_row].fillna("").tolist()
        logging.info("Cleaning the df.columns")
        df=clean_columns(df)

        # drop rows to header and get data below it
        df=df.dropna(how="all").reset_index(drop=True)

        # try to coerc numeric columns where it make sense
        for col in df.columns:
            try:    
                logging.info("Now droping the uncessary columns")
                sample=df[col].astype(str).str.strip().replace({"nan":""})
                non_empty=sample[sample!=""]
                if len(non_empty)==0:
                    continue
                num_like=non_empty.str.match(r'^-?\d+(\.\d+)?$').sum()
                if num_like>=max(1,int(0.4 * len(non_empty))):
                    df[col]=pd.to_numeric(df[col],errors="coerce")
            except Exception as e:
                continue
            
        logging.info("returning the df after cleaning it")
        
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
            logging.info("Now returning the index of the header where the excel file is started")
        return 0
    except Exception as e:
        raise CustomException(e,sys)
    
def promote_first_row_if_header(df:pd.DataFrame)-> pd.DataFrame:
    """ 
    If the first row looks like real header (not 'unnamed',mostly text),
    promote it to be the Dataframe's column names.
    """
    if df.empty:
        return df
    
    first_row=df.iloc[0].astype(str).str.strip()
    non_empty=[c for c in first_row if c and not c.lower().startswith("unnamed")]

    # heuirestic: if at least half the columns are non empyt text-> treat it as a header
    if len(non_empty) >=(len(df.columns)//2):
        df=df[1:].reset_index(drop=True)
        df.columns=first_row.tolist()
    return df

def clean_columns(df:pd.DataFrame)->pd.DataFrame:
    #drop fully empty columns
    has_unnamed=any(isinstance(c,str)and c.strip().lower().startswith("unnamed")
                    for c in df.columns
                )
    
    df=df.dropna(axis=1,how="all")
    
    
    if not has_unnamed:
        return df
    else:
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
    file_path = r"examples\Portfolio Allocation Data.xlsx"
    
    aa=excel_reader(path=file_path,sheet=None)
    
    print(aa)

    
