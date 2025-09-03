import pandas as pd
import os
file_path=os.path.join("Examples","Portfolio Allocation Data.xlsx")
df_raw=pd.read_excel(file_path)
print(df_raw)