""" Chart Type Detection Model
    Automatically determine the appropriate chart type based on data characteristics.
"""

import pandas as pd 
from typing import Tuple,Optional,List


def detect_chart_type(df:pd.DataFrame,x_col:Optional[str]=None,y_cols:Optional[str]=None)->Tuple[str,dict]:
    """
    inteliigent detect the best chart type for the data"""
    if df.empty or df.shape[0]<2:
        return  None,{}
    

    # Get Nummerics and categorical columns
    numeric_cols=df.select_dtypes(include=['int64','float64']).columns.tolist()
    categotical_cols=df.select_dtypes(include=['object','string']).columns.tolist()

    # CAse 1 : Pie Chart - Part-to-whole with one categgory and one value 
    #Example:Assests Allocation
    if len(categotical_cols)>=1 and len(numeric_cols)>=1 and df.shape[0]<=10:
        cat_col=categotical_cols[0]
        val_col=numeric_cols[0]

    # Check if data represents parts of a whole
    # KeyWord
    keywords=['allocation','distribution','composition','breakdown','share','percentage','porfolio','sector','country','category','region']
    column_text=''.join(df.columns).lower()
    if any(keyword in column_text for keyword in keywords):
        return 'pie',{
            'category_col':cat_col,
            'value_col':val_col,
            'title':f'{val_col} by {cat_col}'
        }
    
    ## CAse 2:Line Chart - Time series or sequential data
    #Example:Quaterly Revenue,Monthly sales
    if len(categotical_cols)>=1 and len(numeric_cols)>=1:
        cat_col=categotical_cols[0]

        #Check if colunm contains time.sequnece indiciators
        time_keywords=['quater','month','year','week','day','date','q1','q2','q3','q4',
                       'jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec',
                       'time','period']
        
        first_col_text=str(cat_col).lower()
        sample_values=df[cat_col].astype(str).str.lower().str.cat(sep=' ')

        if any (keyword in first_col_text for keyword in time_keywords) or any(keyword in sample_values for keyword in time_keywords):
            return 'line',{
                'x_col':cat_col,
                'y_cols':numeric_cols,
                'title':f'Trend of {", ".join(numeric_cols)} over {cat_col}'
            }
        


    # Case 3: Bar Chart - Compare categories across multiple values
    if len(categotical_cols)>=1 and len(numeric_cols)==1:
        cat_col=categotical_cols[0]
        val_col=numeric_cols[0]
        if df.shape[0]>=3:
            return 'bar',{
                'x_col':cat_col,
                'y_col':val_col,
                'title':f'{val_col} by {cat_col}'
            }
        
    # Case 4: COLUMN Chart - Compare categories across multiple values
    if len(categotical_cols)>=1 and len(numeric_cols)>1 and df.shape[0]<=12:
        return 'column',{
            'x_col':categotical_cols[0],
            'y_cols':numeric_cols,
            'title':f'Comparison of {", ".join(numeric_cols)} by {categotical_cols[0]}'
        }
    
    #Case 5: Scatter Plot - Relationship between two numeric variables
    if len(numeric_cols)>=2:
        return 'scatter',{
            'x_col':numeric_cols[0],
            'y_cols':numeric_cols[1],
            'title':f'{numeric_cols[1]} vs {numeric_cols[0]}'
        }
    
    # Default: Column Chart if we have any data
    if len(numeric_cols)>0:
        return 'column',{
            'x_col':categotical_cols[0] if categotical_cols else None,
            'y_cols':numeric_cols[:3],
            'title':'Data Overview'
        }
    return None,{}



def should_create_chart(df:pd.DataFrame,min_rows:int=2,max_rows:int=50)->bool:
    """Determine if a chart should be created based on data size"""
    if df is None or df.empty:
        return False
    
    if df.shape[0]<min_rows or df.shape[0]>max_rows:
        return False
    
    numeric_cols=df.select_dtypes(include=['int64','float64']).columns
    if len(numeric_cols)==0:
        return False
    has_data=df[numeric_cols].notna().any().any()
    return has_data