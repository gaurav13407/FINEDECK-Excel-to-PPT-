import os
import re
import logging
from typing import Optional, Tuple, Dict, Any

import pandas as pd

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)


def excel_to_csv(excel_path: str, csv_path: Optional[str] = None, sheet_name: Optional[str] = None) -> str:
    """
    Read an Excel file and write the first (or specified) sheet to CSV.

    Returns the path to the created CSV.
    """
    if csv_path is None:
        csv_path = os.path.splitext(excel_path)[0] + ".csv"

    try:
        
        df = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl")
    except Exception:
        
        df = pd.read_excel(excel_path, sheet_name=sheet_name)

    if isinstance(df, dict):
        # multiple sheets returned; pick first
        first_sheet = list(df.keys())[0]
        df = df[first_sheet]

    if df.empty:
        logger.warning("Excel file loaded but contains no data: %s", excel_path)

    os.makedirs(os.path.dirname(csv_path) or ".", exist_ok=True)
    df.to_csv(csv_path, index=False)

    logger.info("Wrote CSV: %s", csv_path)
    return csv_path


def clean_csv(csv_path: str, cleaned_path: Optional[str] = None, *, require_columns: Tuple[str, ...] = ("date", "amount")) -> str:
    """
    Clean a CSV produced from financial data.

    Steps:
    - Normalize column names
    - Map common alternate column names to canonical names
    - Parse dates robustly
    - Clean amount strings (currency symbols, thousand separators, parentheses → negative)
    - Drop rows missing required columns after parsing
    - Drop duplicates
    - Add `month` and `year` columns when date exists

    Returns path to cleaned CSV.
    """
    if cleaned_path is None:
        cleaned_path = os.path.splitext(csv_path)[0] + "_cleaned.csv"

    if not os.path.exists(csv_path):
        raise FileNotFoundError(f"CSV not found: {csv_path}")

    df = pd.read_csv(csv_path)

    # normalize column names
    df.columns = [str(c).strip().lower() for c in df.columns]

    rename_map = {
        "transaction date": "date",
        "posted date": "date",
        "date": "date",
        "description": "description",
        "desc": "description",
        "narration": "description",
        "amount": "amount",
        "value": "amount",
        "category": "category",
        "type": "category",
    }

    df = df.rename(columns={c: rename_map.get(c, c) for c in df.columns})

    # Parse dates if present
    if "date" in df.columns:
        # try to coerce with pandas; provide common formats fallback
        df["date"] = pd.to_datetime(df["date"], errors="coerce")

    # Clean amounts if present
    if "amount" in df.columns:
        s = df["amount"].astype(str)
        # normalize whitespace and NBSP
        s = s.str.replace(r"\u00A0", " ", regex=True).str.strip()
        # remove currency symbols/letters, keep digits, dot, comma, parentheses, minus
        s = s.str.replace(r"[^0-9\.,\-()]+", "", regex=True)
        # parentheses indicate negative values: (1,234.56) -> -1,234.56
        s = s.str.replace(r"^\((.*)\)$", r"-\1", regex=True)
        # remove thousand separators (commas)
        s = s.str.replace(r",", "", regex=True)
        # final numeric conversion
        df["amount"] = pd.to_numeric(s, errors="coerce")

    # Ensure required columns exist
    missing = [c for c in require_columns if c not in df.columns]
    if missing:
        raise ValueError(f"Required columns missing after normalization: {missing}")

    # Drop rows where required columns are NaN after parsing
    before_rows = len(df)
    df = df.dropna(subset=list(require_columns))
    dropped_rows = before_rows - len(df)

    df = df.drop_duplicates()

    if "date" in df.columns:
        df["month"] = df["date"].dt.month
        df["year"] = df["date"].dt.year

    os.makedirs(os.path.dirname(cleaned_path) or ".", exist_ok=True)
    # atomic write: write to temp then move
    tmp_path = cleaned_path + ".tmp"
    df.to_csv(tmp_path, index=False)
    os.replace(tmp_path, cleaned_path)

    logger.info("Cleaned CSV written: %s (dropped %d rows)", cleaned_path, dropped_rows)
    return cleaned_path
