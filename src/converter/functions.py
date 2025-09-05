import pandas as pd
import openpyxl
from openpyxl.styles import Font,Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from math import sqrt
from typing import Optional,Dict,Any
import numpy as np
import re

##------------Utilities--------------
""" 
Later Gonna add Risk metrics(Sortino,Max Drawdown,Beta vs Index)
Vaulation model(PE ratios,multiples).
Scenario comparision(Bull,Base,Bear in separate sheets)
"""
COMMON_PNL_NAMES=["pnl","p&l","profit","profit_and_loss","pl"]
COMMON_BALANCE_NAMES=["balance","bs","balance_sheet"]
COMMON_RETURNS_NAMES=["returns","return","performance","return_series","price","prices","nav"]

def _noramlize_name(name:str)->str:
    return re.sub(r"[^a-z0-9]","",name.lower())

def _sheet_matches(name:str, keywords:list)->bool:
    n=_noramlize_name(name)
    for k in keywords:
        if _noramlize_name(k) is n:
            return True
    return False

#--------------------Detection / Extraction--------------
def find_best_sheet_for_pnl(xls:pd.ExcelFile)->Optional[str]:
    for sheet in xls.sheet_names:
        if _sheet_matches(sheet,COMMON_PNL_NAMES):
            return sheet
        
    ##fallback:try sheets with typical Pnl -Like columns
    for sheet in xls.sheet_names:
        df=pd.read_excel(xls,sheet_name=sheet, nrows=5)
        cols=" ".join([c.lower() for c in df.columns.astype(str)])
        if "revenu" in cols or "profit" in cols or "gross" in cols:
            return sheet
        
    return None

def find_best_sheet_for_return(xls:pd.ExcelFile)->Optional[str]:
    for sheet in xls.sheet_names:
        if _sheet_matches(sheet,COMMON_RETURNS_NAMES):
            return sheet
        
    ## fallback:sheet that has a data-like column + numeric time series
    for sheet in xls.sheet_names:
        df=pd.read_excel(xls,sheet_name=sheet,nrows=10)
        cols=df.columns.astype(str)
        if any(re.search(r"date|day|month|year",c.lower()) for c in cols):
            # check numeric columns exists
            if df.select_dtypes(include=[np.number]).shape[1] >=1:
                return sheet
    return None

#----------------Metrics--------------
def total_revenu_from_pnl(df:pd.DataFrame)->Optional[float]:
    #try common columns name
    for cand in ["revenu","total revenue","sales"]:
        for c in df.columns:
            if cand in str(c).lower():
                try:
                    return float(df[c].replace({np.nan:0}).sum())
                except Exception:
                    try:
                        return float(pd.to_numeric(df[c],error="coerce").sum())
                    except Exception:
                        pass
    ## Fallback:largest numeric column sum
    nums_cols=df.select_dtypes(include=[np.number]).columns
    if len(nums_cols):
        s=df[nums_cols].sum(axis=0)
        # Chose column with max absolute sum
        col=s.abs().idxmax()
        return float(s[col])
    return None


## Be aware of this function might be a bug here
def total_profit_from_pnl(df:pd.DataFrame)-> Optional[float]:
    for cand in ["net income","profit","netprofit","net profit","net income(loss)"]:
        for c in df.columns:
            if cand in str(c).lower():
                try:
                    return float(pd.to_numeric(df[c],erros="coerc").sum())
                except Exception:
                    pass
    
    # fallback use last row if labeled 'Total' or last numeric column total
    if df.shape[0] >0:
        last_row=df.iloc[-1]
        numerics=pd.to_numeric(last_row.select_dtypes(inculde=[np.number]), errors="coerce")
        if not numerics.dropna().empty:
            return float(numerics.dropna().sum())
        
    return None

def compute_sharpe(returns:pd.Series,returns_freq: int=252)-> Optional[float]:
    """Compute annualized Sharper ratio(assume return are simple return).
    return_freq:period pre year (252 daily, 12 months,52 weekly)
    """
    r=pd.to_numeric(returns.dropna(), errors="coerce")
    if r.empty or r.std()==0:
        return None
    mean=r.mean()
    std=r.std(ddof=1)
    #annualize
    sharper=(mean * returns_freq)/(std*sqrt(returns_freq))
    return float(sharper)

def compute_cagr_from_price(price_series:pd.Series)-> Optional[float]:
    """ Compute CAGR from price/time series(assume price sereis orderd by time)."""
    s=price_series.dropna().astype(float)
    if len(s)<2:
        return None
    start=s.iloc[0]
    end=s.iloc[-1]
    periods=len(s) -1
    # we'll not assume exact frequency (user can override), so compute simple CAGR per period:
    try:
        cagr=(end/start) ** (1.0/periods)-1
        return float(cagr)
    
    except Exception:
        return None
    
#------------------Summary Writer------------------
def write_summary_sheet(wb_path:str,metrics:Dict[str,Any],table_preview:Optional[pd.DataFrame]=None) -> str:
    wb=openpyxl.load_workbook(wb_path)
    if "Summary" in wb.sheetnames:
        del wb["summary"]

    ws=wb.create_sheet("Summary",0)
    ws["A1"]="Auto-generated Finance Summary"
    ws["A1"].font=Font(bold=True,size=14)
    ws["A1"].alignment=Alignment(horizontal="left")

    row=3
    for k,v in metrics.items():
        ws.cell(row=row,column=1,value=str(k))
        ws.cell(row=row,column=2,value=(v if v is not None else "N/A"))
        row+=1

    # If a small table preview is provided ,paste it below metrices
    if table_preview is not None and not table_preview.empty:
        start_row=row+1
        ws.cell(row=start_row,column=1,value="Sample Table Preview")
        ws.cell(row=start_row,column=1).font=Font(bold=True)
        for r_idx, r in enumerate(dataframe_to_rows(table_preview.head(10),index=False,header=True),start=start_row+1):
            for c_idx, val in enumerate(r,start=1):
                ws.cell(row=r_idx,column=c_idx,value=val)

    wb.save(wb_path)
    return wb_path

#----------------Main Orchestraor--------
def create_summary_from_excel(file_path:str,return_freq:int =252,prefer_sheet:Optional[str]=None)->str:
    """ 
    Inspects the excel file,detect PnL/return sheets,computes metrices and write a Summary sheet metrices and writes a Summary sheet.
    return_freq:period per year to annualize Sharpe(252 daily, 12 monthly....)
    """
    xls=pd.ExcelFile(file_path)
    metrics={}

    # Detect PnL
    pnl_sheet=prefer_sheet or find_best_sheet_for_pnl(xls)