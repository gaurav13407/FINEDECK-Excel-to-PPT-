"""
Quick runner to test `excel_to_csv` and `clean_csv` functions.
Run from repository root:

    python scripts\run_file_converter_test.py

"""
import sys
from pathlib import Path
import traceback

# ensure project src is importable
project_root = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(project_root))

try:
    from src.Power_BI_Converter.upload.file_converter import excel_to_csv, clean_csv
except Exception as e:
    print('Import failed:', e)
    traceback.print_exc()
    raise SystemExit(1)

TEST_FILE = Path('examples') / 'Sample_pnl.xlsx'

if not TEST_FILE.exists():
    print('Test Excel not found:', TEST_FILE)
    raise SystemExit(1)

try:
    print('Converting Excel -> CSV...')
    csv = excel_to_csv(str(TEST_FILE))
    print('CSV created at:', csv)

    print('Cleaning CSV...')
    cleaned = clean_csv(csv)
    print('Cleaned CSV created at:', cleaned)

    import pandas as pd
    df = pd.read_csv(cleaned)
    print('Cleaned rows:', len(df))
    print(df.head().to_string())

except Exception as e:
    print('Error during test:')
    traceback.print_exc()
    raise SystemExit(1)
