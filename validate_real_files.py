"""Test validation against real files in input_data/."""

import pandas as pd
from toll_processor.processor import process_toll_files

linkt_file = "input_data/History_9211592178_20260909123318.xls"
bookings_file = "input_data/Booking-list.xls"
master_file = "input_data/CURRENT TOLL SHEET (TESTING).xlsx"

print("=" * 80)
print("TEST RUNNING PIPELINE WITH REAL EXCEL FILES")
print("=" * 80)

result = process_toll_files(
    linkt_file=linkt_file,
    bookings_365_file=bookings_file,
    master_file=master_file,
    date_received="2026-09-08",
    admin_fee=5.55,
)

stats = result.summary_stats
print("\n>>> SUMMARY RESULTS ON REAL DATA:")
for k, v in stats.items():
    print(f"  {k:28s}: {v}")

print(f"\nTotal rows in Updated Master : {len(result.master_df)} (Original had 2,351)")
print(f"New Tolls Matched             : {len(result.new_tolls_df)}")
print(f"Duplicates Detected           : {len(result.duplicates_df)}")
print(f"Already Paid Detected         : {len(result.already_paid_df)}")
print(f"Unmatched Records             : {len(result.unmatched_df)}")

print("\nSample 5 New Matched Tolls:")
if not result.new_tolls_df.empty:
    cols_to_show = [c for c in ["REGO", "Date", "Time", "Driver", "Toll Amount", "Admin Fee", "TOTAL AMOUNT ", "Hire Date", "Return Date"] if c in result.new_tolls_df.columns]
    print(result.new_tolls_df[cols_to_show].head(5))

print("\nCustomer Invoicing Summary (Top 5):")
if not result.customer_summary_df.empty:
    print(result.customer_summary_df.head(5))

print("\nSample Unmatched Records (Top 5):")
if not result.unmatched_df.empty:
    print(result.unmatched_df[["rego", "start_datetime", "toll_amount", "status", "reason"]].head(5))

print("\nSample Duplicate Records (Top 5):")
if not result.duplicates_df.empty:
    print(result.duplicates_df[["rego", "start_datetime", "toll_amount", "reason"]].head(5))

print("\nSample Already Paid Records (Top 5):")
if not result.already_paid_df.empty:
    print(result.already_paid_df[["rego", "start_datetime", "toll_amount", "reason"]].head(5))
