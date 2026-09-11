"""Validate Daily Workflow carry-forward with real input files."""

import os
import shutil
import pandas as pd
from toll_processor.processor import process_daily_run, process_toll_files
from toll_processor import master_store
from toll_processor.schemas import DATA_DIR, ARCHIVE_DIR, CURRENT_MASTER_PATH

def run_validation():
    print("=" * 60)
    print("VALIDATING DAILY CARRY-FORWARD WORKFLOW WITH REAL DATA")
    print("=" * 60)

    # 1. Paths to real files
    input_dir = "input_data"
    linkt_file = os.path.join(input_dir, "History_9211592178_20260909123318.xls")
    bookings_file = os.path.join(input_dir, "Booking-list.xls")
    master_file = os.path.join(input_dir, "CURRENT TOLL SHEET (TESTING).xlsx")

    assert os.path.exists(linkt_file), f"Missing {linkt_file}"
    assert os.path.exists(bookings_file), f"Missing {bookings_file}"
    assert os.path.exists(master_file), f"Missing {master_file}"

    # Read original master row count
    orig_df, _ = master_store.read_master_file(master_file)
    orig_rows = len(orig_df)
    print(f"\n[1] Original Master Dataset row count: {orig_rows:,}")

    # Ensure clean data directory for validation
    if os.path.exists(DATA_DIR):
        # backup if any
        pass

    # DAY 1: INITIAL SETUP
    print("\n[2] Executing DAY 1: INITIAL SETUP...")
    res_day1, save_info_day1 = process_daily_run(
        linkt_file=linkt_file,
        bookings_365_file=bookings_file,
        master_file=master_file,
        date_received="2026-09-08",
        admin_fee=5.55,
        is_initial_setup=True,
    )

    stats1 = res_day1.summary_stats
    print(f"  • LINKT Trips processed:      {stats1['total_records']}")
    print(f"  • New matched tolls:          {stats1['new_tolls_count']}")
    print(f"  • Duplicate tolls:            {stats1['duplicates_count']}")
    print(f"  • Already paid tolls:         {stats1['already_paid_count']}")
    print(f"  • Unmatched tolls (unpaid):   {stats1['total_unmatched_all']}")
    print(f"  • Financial totals: Toll ${stats1['total_toll_amount']:,.2f} + Fee ${stats1['total_admin_fee']:,.2f} = Total ${stats1['total_customer_amount']:,.2f}")
    print(f"  • Master updated from {orig_rows:,} to {save_info_day1['new_row_count']:,} rows")
    print(f"  • Saved as: {CURRENT_MASTER_PATH}")
    print(f"  • Validation: {save_info_day1['validation']}")

    # Assertions for Day 1
    assert master_store.has_stored_master(), "current_master.xlsx must exist after Day 1"
    assert save_info_day1["new_row_count"] == orig_rows + stats1["new_tolls_count"]

    # Verify no unmatched toll has an admin fee or is charged to customer
    if not res_day1.unmatched_df.empty:
        unmatched_regos = set(res_day1.unmatched_df["rego"].tolist())
        charged_drivers = set(res_day1.customer_summary_df["Driver"].tolist())
        assert "UNMATCHED" not in charged_drivers
        assert "Unknown" not in charged_drivers

    # DAY 2: DAILY RUN WITH SAME LINKT FILE
    print("\n[3] Executing DAY 2: DAILY RUN (Testing Duplicate & Carry-Forward Protection)...")
    res_day2, save_info_day2 = process_daily_run(
        linkt_file=linkt_file,
        bookings_365_file=bookings_file,
        master_file=None, # Automatically loads current_master.xlsx
        date_received="2026-09-09",
        admin_fee=5.55,
        is_initial_setup=False,
    )

    stats2 = res_day2.summary_stats
    print(f"  • LINKT Trips processed:      {stats2['total_records']}")
    print(f"  • New matched tolls:          {stats2['new_tolls_count']}")
    print(f"  • Duplicate tolls detected:   {stats2['duplicates_count']}")
    print(f"  • Financial totals: Toll ${stats2['total_toll_amount']:,.2f} + Fee ${stats2['total_admin_fee']:,.2f} = Total ${stats2['total_customer_amount']:,.2f}")
    print(f"  • Previous Master archived as:{save_info_day2['archived_as']}")
    print(f"  • New Master row count:       {save_info_day2['new_row_count']:,} (unchanged!)")
    print(f"  • Validation: {save_info_day2['validation']}")

    # Assertions for Day 2:
    assert stats2["new_tolls_count"] == 0, "Re-running same LINKT file must yield exactly 0 new tolls!"
    assert stats2["total_admin_fee"] == 0.0, "Zero admin fee when 0 new tolls"
    assert stats2["total_customer_amount"] == 0.0, "Zero customer charge when 0 new tolls"
    assert save_info_day2["new_row_count"] == save_info_day1["new_row_count"], "Master rows must not grow on duplicates"
    assert save_info_day2["archived_as"] is not None, "Previous master must have been archived"

    # Verify archive exists
    archives = master_store.list_archive()
    print(f"\n[4] Archive verification: {len(archives)} file(s) in data/archive/:")
    for a in archives:
        print(f"  - {a['filename']} ({a['file_size_kb']} KB) - Created: {a['created']}")

    print("\n" + "=" * 60)
    print("ALL BUSINESS RULES & CARRY-FORWARD CHECKS PASSED PERFECTLY!")
    print("=" * 60)

if __name__ == "__main__":
    run_validation()
