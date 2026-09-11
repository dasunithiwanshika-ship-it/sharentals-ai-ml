"""Comprehensive analysis of real input files."""

import pandas as pd
import numpy as np

print("=" * 80)
print("COMPREHENSIVE DATA PROFILING")
print("=" * 80)

# 1. 365 Bookings
df_365 = pd.read_excel("input_data/Booking-list.xls", engine="calamine")
print(f"365 BOOKINGS: {len(df_365)} rows")
print("Columns:", list(df_365.columns))
print("\nSample Finish Date values in 365:")
print(df_365["Finish Date"].value_counts().head(10))

# 2. LINKT History
df_linkt = pd.read_excel("input_data/History_9211592178_20260909123318.xls")
print(f"\nLINKT REPORT: {len(df_linkt)} rows")
print("Columns:", list(df_linkt.columns))
print("\nUnique Types in LINKT:")
print(df_linkt["Type"].value_counts())
print("\nNon-Trips Rows preview:")
print(df_linkt[df_linkt["Type"] != "Trips"][["Start Date", "Type", "Details", "LPN", "Amount"]].head(10))

# 3. Master Dataset
xl_master = pd.ExcelFile("input_data/CURRENT TOLL SHEET (TESTING).xlsx")
print(f"\nMASTER DATASET SHEETS: {xl_master.sheet_names}")
for s in xl_master.sheet_names:
    df_s = pd.read_excel(xl_master, sheet_name=s)
    print(f"  Sheet '{s}': {df_s.shape[0]} rows, {df_s.shape[1]} cols")
    print(f"    Columns: {list(df_s.columns[:10])}...")
    if s == "Paste_eToll_Data":
        print("    Paste_eToll_Data full columns:")
        print(list(df_s.columns))
        print("    Paste_eToll_Data Date Received value counts (top 10):")
        print(df_s["Date Received"].value_counts(dropna=False).head(10))
        print("    Sample 5 rows:")
        print(df_s.head(5)[["Unique Key", "Date Received", "REGO", "Toll Amount", "Admin Fee", "TOTAL AMOUNT ", "Driver", "Payment Status"]])
