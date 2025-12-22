import os
import sys
from pathlib import Path
import shutil
import pytest

# Make the repository root importable when running this file directly
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from src.Power_BI_Converter.upload.file_converter import excel_to_csv, clean_csv

EXAMPLE = Path("examples") / "Sample_pnl.xlsx"


def setup_module(module):
    # ensure examples file exists for test
    assert EXAMPLE.exists(), f"Test example not found: {EXAMPLE}"


def test_excel_to_csv_and_clean(tmp_path):
    # create workspace in tmp
    work_dir = tmp_path / "work"
    work_dir.mkdir()

    # copy example into tmp to avoid modifying originals
    excel_copy = work_dir / EXAMPLE.name
    shutil.copy(EXAMPLE, excel_copy)

    csv_path = work_dir / (excel_copy.stem + ".csv")
    cleaned_path = work_dir / (excel_copy.stem + "_cleaned.csv")

    # Run conversion
    created_csv = excel_to_csv(str(excel_copy), csv_path=str(csv_path))
    assert Path(created_csv).exists(), "CSV was not created"

    # Run cleaning
    created_cleaned = clean_csv(str(csv_path), cleaned_path=str(cleaned_path))
    assert Path(created_cleaned).exists(), "Cleaned CSV was not created"

    # Basic sanity checks on cleaned file
    import pandas as pd
    df = pd.read_csv(created_cleaned)

    # Required columns should be present
    assert "date" in df.columns
    assert "amount" in df.columns

    # amount should be numeric (no string types)
    assert pd.api.types.is_numeric_dtype(df["amount"]), "Amount column not numeric"

    # date column should be parseable to datetime when reloaded
    try:
        pd.to_datetime(df["date"])  # if this doesn't raise, parseable
    except Exception:
        pytest.fail("Date column not parseable")

    # Cleanup
    for p in [created_csv, created_cleaned]:
        if os.path.exists(p):
            os.remove(p)
