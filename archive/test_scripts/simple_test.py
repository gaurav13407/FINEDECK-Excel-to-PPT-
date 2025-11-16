import sys
sys.path.insert(0, r'c:\Users\gaura\OneDrive\Desktop\Big Projects\FinDeck(Excel to PPT Project)\src')

from converter.excel_reader import excel_reader  
from converter.ppt_writer import df_to_ppt

# Simple test with your backend converter
excel_file = r'c:\Users\gaura\OneDrive\Desktop\Big Projects\FinDeck(Excel to PPT Project)\DV+Sales+Data.xlsx'

print("Reading Excel...")
df = excel_reader(excel_file, sheet=0)
print(f"Loaded: {len(df)} rows")

print("\nGenerating PPT...")
df_to_ppt(
    df=df,
    out_path='Simple_Backend_Test.pptx',
    title="Sales Data Report",
    subtitle="Generated from Backend",
    title_col=None,
    mode="table",
    limit=100  # Limit to 100 rows for faster testing
)

print("✅ Done! Check Simple_Backend_Test.pptx")
