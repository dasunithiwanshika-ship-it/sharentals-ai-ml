"""Script to inspect the real input Excel files in input_data/."""

import os
import sys
import pandas as pd
import openpyxl

input_dir = "input_data"
files = [
    "History_9211592178_20260909123318.xls",
    "Booking-list.xls",
    "CURRENT TOLL SHEET (TESTING).xlsx",
]

print("=" * 80)
print("INSPECTING REAL EXCEL FILES IN input_data/")
print("=" * 80)

def try_read(filepath):
    """Try reading via pandas or openpyxl or html if .xls is HTML table."""
    try:
        # Try default
        xl = pd.ExcelFile(filepath)
        return xl, "ExcelFile"
    except Exception as e1:
        # Check if it's HTML table named .xls
        try:
            dfs = pd.read_html(filepath)
            return dfs, "HTML_TABLE"
        except Exception as e2:
            return None, f"Error: {e1} | HTML Error: {e2}"

for fname in files:
    fpath = os.path.join(input_dir, fname)
    if not os.path.exists(fpath):
        print(f"File not found: {fpath}")
        continue
    
    print(f"\n>>> FILE: {fname}")
    print(f"File Size: {os.path.getsize(fpath):,} bytes")
    
    res, mode = try_read(fpath)
    print(f"Reader format detected: {mode}")
    
    if mode == "ExcelFile":
        print(f"Sheet Names: {res.sheet_names}")
        for sname in res.sheet_names:
            df_raw = pd.read_excel(res, sheet_name=sname, header=None)
            print(f"  --- Sheet: '{sname}' (Total Rows: {len(df_raw)}, Total Cols: {df_raw.shape[1]}) ---")
            print("  First 10 rows preview:")
            for r_idx in range(min(10, len(df_raw))):
                vals = [str(v) if pd.notna(v) else "" for v in df_raw.iloc[r_idx].values]
                # Filter trailing blanks
                while vals and vals[-1] == "":
                    vals.pop()
                print(f"    Row {r_idx:2d}: {vals[:10]}")
    elif mode == "HTML_TABLE":
        print(f"Detected HTML Table in .xls! Total tables: {len(res)}")
        df_raw = res[0]
        print(f"  Table 0 (Total Rows: {len(df_raw)}, Total Cols: {df_raw.shape[1]})")
        print("  First 10 rows preview:")
        for r_idx in range(min(10, len(df_raw))):
            vals = [str(v) if pd.notna(v) else "" for v in df_raw.iloc[r_idx].values]
            while vals and vals[-1] == "":
                vals.pop()
            print(f"    Row {r_idx:2d}: {vals[:10]}")
    else:
        print(f"Failed to read: {mode}")
