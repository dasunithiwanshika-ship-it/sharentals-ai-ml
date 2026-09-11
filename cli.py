"""Command Line Interface for Toll Processing Automation System."""

import argparse
import os
import sys
import pandas as pd
from toll_processor.processor import process_toll_files
from toll_processor.sample_generator import generate_sample_datasets


def main():
    parser = argparse.ArgumentParser(
        description="Toll Processing Automation System - LINKT & 365 Matching Engine"
    )
    parser.add_argument(
        "--linkt",
        "-l",
        type=str,
        help="Path to LINKT Toll Report Excel file",
    )
    parser.add_argument(
        "--bookings",
        "-b",
        type=str,
        help="Path to 365 Rental Booking Export Excel file",
    )
    parser.add_argument(
        "--master",
        "-m",
        type=str,
        default=None,
        help="Path to Existing Master Dataset Excel file (optional)",
    )
    parser.add_argument(
        "--date-received",
        "-d",
        type=str,
        default="2026-09-08",
        help="Date Received for new imports (format: YYYY-MM-DD or DD/MM/YYYY, default: 2026-09-08)",
    )
    parser.add_argument(
        "--admin-fee",
        "-f",
        type=float,
        default=5.55,
        help="Admin Fee to apply per new toll (default: $5.55)",
    )
    parser.add_argument(
        "--output-dir",
        "-o",
        type=str,
        default="output",
        help="Output directory for generated files (default: output)",
    )
    parser.add_argument(
        "--generate-samples",
        action="store_true",
        help="Generate sample datasets in sample_data/ directory and exit",
    )

    args = parser.parse_args()

    if args.generate_samples:
        print("Generating sample datasets...")
        l_path, b_path, m_path = generate_sample_datasets("sample_data")
        print(f"  - LINKT Sample: {l_path}")
        print(f"  - 365 Bookings Sample: {b_path}")
        print(f"  - Master Dataset Sample: {m_path}")
        print("Sample generation complete.")
        return 0

    if not args.linkt or not args.bookings:
        print("Error: --linkt and --bookings are required. Or use --generate-samples.")
        parser.print_help()
        return 1

    if not os.path.exists(args.linkt):
        print(f"Error: LINKT file not found: {args.linkt}")
        return 1

    if not os.path.exists(args.bookings):
        print(f"Error: 365 Bookings file not found: {args.bookings}")
        return 1

    if args.master and not os.path.exists(args.master):
        print(f"Error: Master file not found: {args.master}")
        return 1

    os.makedirs(args.output_dir, exist_ok=True)

    print("=" * 65)
    print("TOLL PROCESSING AUTOMATION SYSTEM")
    print("=" * 65)
    print(f"LINKT File     : {args.linkt}")
    print(f"365 Bookings   : {args.bookings}")
    print(f"Master Dataset : {args.master if args.master else '(None - New Master will be created)'}")
    print(f"Date Received  : {args.date_received}")
    print(f"Admin Fee      : ${args.admin_fee:.2f}")
    print("-" * 65)
    print("Processing...")

    result = process_toll_files(
        linkt_file=args.linkt,
        bookings_365_file=args.bookings,
        master_file=args.master,
        date_received=args.date_received,
        admin_fee=args.admin_fee,
    )

    stats = result.summary_stats
    file_suffix = stats["file_date_suffix"]
    output_master_file = os.path.join(
        args.output_dir, f"Master_Toll_Updated_{file_suffix}.xlsx"
    )

    with open(output_master_file, "wb") as f:
        f.write(result.excel_bytes)

    print("\n" + "=" * 65)
    print("PROCESSING SUMMARY")
    print("=" * 65)
    print(f"Existing Master Records (Before) : {stats.get('master_rows_before_clean', 0)}")
    print(f"  ├── Duplicate Rows Cleaned     : {stats.get('master_dupes_removed', 0)}")
    print(f"  └── Clean Historical Records   : {stats.get('master_rows_after_clean', 0)}")
    print("-" * 65)
    print(f"Total LINKT Records Processed    : {stats['total_linkt_records']}")
    print(f"  ├── New Matched Tolls          : {stats['new_tolls_count']}")
    print(f"  ├── LINKT Duplicates Skipped   : {stats['duplicates_count']}")
    print(f"  ├── Already Paid Tolls         : {stats['already_paid_count']}")
    print(f"  └── Unmatched / Multi-Match    : {stats['unmatched_count']}")
    print("-" * 65)
    print(f"Total Toll Amount ($)            : ${stats['total_toll_amount']:,.2f}")
    print(f"Total Admin Fees ($)             : ${stats['total_admin_fee']:,.2f}")
    print(f"Total Customer Charge ($)        : ${stats['total_customer_amount']:,.2f}")
    print("-" * 65)
    print(f"Final Master Row Count           : {stats['final_master_row_count']}")
    print("=" * 65)
    print(f"Output saved to: {output_master_file}")
    print("Sheets included: Paste_eToll_Data, New Tolls, Customer Summary, Duplicates, Unmatched, Already Paid, Summary.")

    return 0


if __name__ == "__main__":
    sys.exit(main())
