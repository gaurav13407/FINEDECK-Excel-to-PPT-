import pandas as pd
import os

def read_excel_1(file_path: str, sheet_name: str = None, all_sheets: bool = False):
    """ 
    Reads an Excel file.
    - If all_sheets=True → returns a dict {sheet_name: DataFrame}
    - Otherwise → returns a single DataFrame
    """
    if all_sheets:
        # Load all sheets
        dfs = pd.read_excel(file_path, sheet_name=None,header=1)
        for name in dfs:
            dfs[name] = dfs[name].loc[:, ~dfs[name].columns.str.contains("^Unnamed")]
            dfs[name] = dfs[name].reset_index(drop=True)
        return  dfs
    else:
        # Load just one sheet
        if sheet_name is None:
            sheet_name = 0  # default → first sheet

        df = pd.read_excel(file_path, sheet_name=sheet_name, header=1)
        df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
        df = df.reset_index(drop=True)
        return df

if __name__ == "__main__":
    file_path=os.path.join("Examples","Portfolio Allocation Data.xlsx")
    ss=pd.read_excel(file_path)
    print(ss)
    df=read_excel_1(file_path)
    type(df)
    print(df)
