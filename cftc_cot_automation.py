#!/usr/bin/env python3
"""
CFTC COT Data Automation
- Downloads latest weekly COT data (Futures Only, XLS format)
- Extracts WTI-PHYSICAL NYMEX (code 067651) or any contract
- Appends new rows to your master Excel file
- Runs weekly via GitHub Actions (Friday after 3:30 PM ET when CFTC publishes)
"""

import os
import sys
import zipfile
import tempfile
import shutil
from datetime import datetime
from pathlib import Path
import requests
import pandas as pd
from openpyxl import load_workbook, Workbook

# ========== CONFIGURATION ==========
MASTER_FILE = Path("data/cot_master.xlsx")      # Your master file (committed to repo)
SHEET_NAME = "WTI_PHYSICAL_NYMEX"               # Sheet name in master
TARGET_CODE = "067651"                          # CFTC_Contract_Market_Code
TARGET_NAME = "WTI-PHYSICAL - NEW YORK MERCANTILE EXCHANGE"

# Columns to keep (adjust as needed)
KEEP_COLS = [
    "Report_Date_as_MM_DD_YYYY",
    "CFTC_Contract_Market_Code",
    "Open_Interest_All",
    "NonComm_Positions_Long_All",
    "NonComm_Positions_Short_All",
    "NonComm_Postions_Spread_All",
    "Comm_Positions_Long_All",
    "Comm_Positions_Short_All",
    "Tot_Rept_Positions_Long_All",
    "Tot_Rept_Positions_Short_All",
    "NonRept_Positions_Long_All",
    "NonRept_Positions_Short_All",
    "Change_in_Open_Interest_All",
    "Change_in_NonComm_Long_All",
    "Change_in_NonComm_Short_All",
    "Pct_of_OI_NonComm_Long_All",
    "Pct_of_OI_NonComm_Short_All",
    "Pct_of_OI_Comm_Long_All",
    "Pct_of_OI_Comm_Short_All",
    "Contract_Units",
]

# CFTC URL pattern (yearly zip files)
BASE_URL = "https://www.cftc.gov/files/dea/history/dea_fut_xls_{year}.zip"
# ====================================


def get_latest_cftc_url() -> str:
    """Determine the latest available CFTC zip URL."""
    now = datetime.now()
    # CFTC publishes weekly data for current year
    # Also check previous year in early January
    for year in [now.year, now.year - 1]:
        url = BASE_URL.format(year=year)
        try:
            resp = requests.head(url, timeout=10)
            if resp.status_code == 200:
                return url
        except Exception:
            pass
    raise RuntimeError("Could not find available CFTC data file")


def download_and_extract(url: str) -> Path:
    """Download zip and extract annual.xls to temp dir."""
    tmpdir = Path(tempfile.mkdtemp(prefix="cftc_"))
    zip_path = tmpdir / "data.zip"
    
    print(f"Downloading {url}...")
    resp = requests.get(url, timeout=60, stream=True)
    resp.raise_for_status()
    zip_path.write_bytes(resp.content)
    
    print("Extracting...")
    with zipfile.ZipFile(zip_path, 'r') as zf:
        # Expect annual.xls inside
        xls_name = "annual.xls"
        if xls_name not in zf.namelist():
            raise FileNotFoundError(f"{xls_name} not in zip: {zf.namelist()}")
        zf.extract(xls_name, tmpdir)
    
    return tmpdir / xls_name


def read_cftc_xls(xls_path: Path) -> pd.DataFrame:
    """Read CFTC .xls file (requires xlrd)."""
    print(f"Reading {xls_path}...")
    df = pd.read_excel(xls_path, engine="xlrd")
    print(f"  Loaded {len(df)} rows, {len(df.columns)} columns")
    return df


def filter_target_contract(df: pd.DataFrame) -> pd.DataFrame:
    """Extract rows for target contract code."""
    # Match by code (primary) and name (fallback)
    mask = (
        df["CFTC_Contract_Market_Code"].astype(str).str.strip() == TARGET_CODE
    ) | (
        df["Market_and_Exchange_Names"].astype(str).str.contains(TARGET_NAME, case=False, na=False)
    )
    result = df.loc[mask, KEEP_COLS].copy()
    result["Report_Date_as_MM_DD_YYYY"] = pd.to_datetime(result["Report_Date_as_MM_DD_YYYY"])
    result = result.sort_values("Report_Date_as_MM_DD_YYYY")
    print(f"  Found {len(result)} rows for {TARGET_NAME} ({TARGET_CODE})")
    return result


def load_master_file() -> pd.DataFrame:
    """Load existing master file or return empty DataFrame."""
    if MASTER_FILE.exists():
        try:
            df = pd.read_excel(MASTER_FILE, sheet_name=SHEET_NAME, engine="openpyxl")
            df["Report_Date_as_MM_DD_YYYY"] = pd.to_datetime(df["Report_Date_as_MM_DD_YYYY"])
            print(f"Loaded master: {len(df)} existing rows")
            return df
        except Exception as e:
            print(f"Warning: Could not read master file: {e}")
    return pd.DataFrame(columns=KEEP_COLS)


def append_new_rows(master_df: pd.DataFrame, new_df: pd.DataFrame) -> pd.DataFrame:
    """Append only new dates not already in master."""
    if master_df.empty:
        return new_df
    
    existing_dates = set(master_df["Report_Date_as_MM_DD_YYYY"].dt.date)
    new_rows = new_df[~new_df["Report_Date_as_MM_DD_YYYY"].dt.date.isin(existing_dates)]
    
    if new_rows.empty:
        print("No new rows to append (all dates already in master)")
        return master_df
    
    print(f"Appending {len(new_rows)} new rows:")
    for _, row in new_rows.iterrows():
        print(f"  + {row['Report_Date_as_MM_DD_YYYY'].date()} OI={row['Open_Interest_All']:,}")
    
    combined = pd.concat([master_df, new_rows], ignore_index=True)
    combined = combined.sort_values("Report_Date_as_MM_DD_YYYY").reset_index(drop=True)
    return combined


def save_master_file(df: pd.DataFrame):
    """Save master file with formatting."""
    MASTER_FILE.parent.mkdir(parents=True, exist_ok=True)
    
    # Write with openpyxl for formatting
    if MASTER_FILE.exists():
        wb = load_workbook(MASTER_FILE)
        if SHEET_NAME in wb.sheetnames:
            del wb[SHEET_NAME]
    else:
        wb = Workbook()
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]
    
    ws = wb.create_sheet(SHEET_NAME)
    
    # Header
    for col_idx, col_name in enumerate(KEEP_COLS, 1):
        ws.cell(row=1, column=col_idx, value=col_name)
    
    # Data
    for row_idx, (_, row) in enumerate(df.iterrows(), 2):
        for col_idx, col_name in enumerate(KEEP_COLS, 1):
            val = row[col_name]
            if pd.isna(val):
                val = None
            elif isinstance(val, (pd.Timestamp, datetime)):
                val = val.date() if isinstance(val, pd.Timestamp) else val
            ws.cell(row=row_idx, column=col_idx, value=val)
    
    # Auto-filter
    ws.auto_filter.ref = f"A1:{chr(64+len(KEEP_COLS))}{len(df)+1}"
    
    wb.save(MASTER_FILE)
    print(f"Saved master: {MASTER_FILE} ({len(df)} rows)")


def cleanup_temp(tmpdir: Path):
    """Remove temp directory."""
    shutil.rmtree(tmpdir, ignore_errors=True)


def main():
    print(f"=== CFTC COT Automation - {datetime.now().isoformat()} ===")
    
    tmpdir = None
    try:
        # 1. Get latest data URL
        url = get_latest_cftc_url()
        print(f"Using: {url}")
        
        # 2. Download & extract
        xls_path = download_and_extract(url)
        tmpdir = xls_path.parent
        
        # 3. Read & filter
        df = read_cftc_xls(xls_path)
        new_data = filter_target_contract(df)
        
        if new_data.empty:
            print("ERROR: No data found for target contract!")
            sys.exit(1)
        
        # 4. Load master & append
        master = load_master_file()
        updated = append_new_rows(master, new_data)
        
        # 5. Save
        save_master_file(updated)
        
        print("=== DONE ===")
        
    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
    finally:
        if tmpdir:
            cleanup_temp(tmpdir)


if __name__ == "__main__":
    main()