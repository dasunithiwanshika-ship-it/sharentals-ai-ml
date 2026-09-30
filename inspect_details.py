"""Detailed inspection script for all 3 actual Excel files."""

import os
import pandas as pd
import openpyxl

def inspect_file_header(fpath):
    with open(fpath, 'rb') as f:
        head = f.read(500)
    print(f"\nRaw first 500 bytes of {fpath}:\n{repr(head)}")

def inspect_booking_file(fpath):
    print("\n" + "=" * 60)
    print(f"INSPECTING 365 BOOKING FILE: {fpath}")
    print("=" * 60)
    # Check if html table
    try:
        dfs = pd.read_html(fpath)
        print(f"pd.read_html succeeded! Number of tables: {len(dfs)}")
        for idx, df in enumerate(dfs):
            print(f"Table {idx} shape: {df.shape}")
            print(f"Columns: {list(df.columns)}")
            print("First 5 rows:")
            print(df.head(5))
            return df
    except Exception as e:
        print(f"pd.read_html failed: {e}")
    
    # Check if xml / excel
    try:
        df = pd.read_xml(fpath)
        print(f"pd.read_xml succeeded! Shape: {df.shape}")
        print(f"Columns: {list(df.columns)}")
        print(df.head(5))
        return df
    except Exception as e:
        print(f"pd.read_xml failed: {e}")

def inspect_linkt_file(fpath):
    print("\n" + "=" * 60)
    print(f"INSPECTING LINKT TOLL FILE: {fpath}")
    print("=" * 60)
    df = pd.read_excel(fpath, sheet_name=0)
    print(f"Shape: {df.shape}")
    print(f"Exact Columns ({len(df.columns)}): {list(df.columns)}")
    print("\nData Types:")
    print(df.dtypes)
    print("\nFirst 10 rows:")
    print(df.head(10))
    print("\nUnique 'Type' values:")
    print(df['Type'].value_counts(dropna=False))
    print("\nAmount statistics:")
    print(df['Amount'].describe())
    print("\nSample LPN values:")
    print(df['LPN'].dropna().unique()[:15])

def inspect_master_file(fpath):
    print("\n" + "=" * 60)
    print(f"INSPECTING EXISTING MASTER DATASET FILE: {fpath}")
    print("=" * 60)
    xl = pd.ExcelFile(fpath)
    print(f"Sheet Names: {xl.sheet_names}")
    for sname in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sname)
        print(f"\n--- Sheet: '{sname}' --- (Shape: {df.shape})")
        print(f"Exact Columns ({len(df.columns)}): {list(df.columns)}")
        if sname == 'Paste_eToll_Data' or sname == 'Sheet1':
            print("\nSample rows:")
            print(df.head(5))
            print("\nUnique 'Date Received' values (sample):")
            if 'Date Received' in df.columns:
                print(df['Date Received'].value_counts(dropna=False).head(10))
            if 'Payment Status' in df.columns:
                print("\nPayment Status counts:")
                print(df['Payment Status'].value_counts(dropna=False).head(10))
            if 'Payment Method' in df.columns:
                print("\nPayment Method counts:")
                print(df['Payment Method'].value_counts(dropna=False).head(10))
            if 'Admin Fee' in df.columns:
                print("\nAdmin Fee unique values:")
                print(df['Admin Fee'].value_counts(dropna=False).head(10))

inspect_file_header("input_data/Booking-list.xls")
inspect_booking_file("input_data/Booking-list.xls")
inspect_linkt_file("input_data/History_9211592178_20260909123318.xls")
inspect_master_file("input_data/CURRENT TOLL SHEET (TESTING).xlsx")
