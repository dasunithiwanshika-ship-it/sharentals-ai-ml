"""End-to-end integration tests for the full Toll Processing pipeline."""

import unittest
import io
import pandas as pd
from toll_processor.sample_generator import generate_sample_datasets
from toll_processor.processor import process_toll_files
from toll_processor.schemas import (
    MASTER_COLUMNS,
    SHEET_MASTER_DATA,
    SHEET_NEW_TOLLS,
    SHEET_DUPLICATES,
    SHEET_UNMATCHED,
    SHEET_ALREADY_PAID,
    SHEET_SUMMARY,
)


class TestPipeline(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.linkt_path, cls.bookings_path, cls.master_path = generate_sample_datasets("test_sample_data")

    def test_full_pipeline_execution(self):
        result = process_toll_files(
            linkt_file=self.linkt_path,
            bookings_365_file=self.bookings_path,
            master_file=self.master_path,
            date_received="2026-09-08",
            admin_fee=5.55,
        )

        stats = result.summary_stats
        # LINKT had 9 total records:
        # - 4 valid matches (ABC123 trip 1, ABC123 trip 2, XYZ789, TOLL99)
        # - 1 overlap multi-match (OVERLAP1) -> Multi-match (1)
        # - 1 unknown rego (NOMATCH99) -> Unmatched (1)
        # - 1 historical duplicate unpaid (HIST001) -> Duplicates
        # - 1 historical duplicate paid (HIST001) -> Already Paid
        # - 1 batch duplicate (ABC123 duplicate) -> Duplicates
        self.assertEqual(stats["total_records"], 9)
        self.assertEqual(stats["new_tolls_count"], 4)
        self.assertEqual(stats["duplicates_count"], 2)
        self.assertEqual(stats["already_paid_count"], 1)
        self.assertEqual(stats["unmatched_count"], 1)
        self.assertEqual(stats["multi_match_count"], 1)
        self.assertEqual(stats["total_unmatched_all"], 2)

        # Toll amounts check
        # Expected new tolls: 7.86 + 3.54 + 9.20 + 4.00 = 24.60
        self.assertAlmostEqual(stats["total_toll_amount"], 24.60, places=2)
        # 4 new tolls * $5.55 = 22.20
        self.assertAlmostEqual(stats["total_admin_fee"], 22.20, places=2)
        # Total customer amount = 24.60 + 22.20 = 46.80
        self.assertAlmostEqual(stats["total_customer_amount"], 46.80, places=2)

        # Check that historical Master data is preserved (original 2 rows + 4 new rows = 6 rows)
        self.assertEqual(len(result.master_df), 6)
        self.assertEqual(result.master_df.iloc[0]["REGO"], "HIST001")
        self.assertEqual(result.master_df.iloc[0]["Admin Fee"], 5.00)

        # Check generated Excel Workbook has all 6 sheets
        wb_file = io.BytesIO(result.excel_bytes)
        excel_reader = pd.ExcelFile(wb_file)
        sheet_names = excel_reader.sheet_names
        self.assertIn(SHEET_MASTER_DATA, sheet_names)
        self.assertIn(SHEET_NEW_TOLLS, sheet_names)
        self.assertIn(SHEET_DUPLICATES, sheet_names)
        self.assertIn(SHEET_UNMATCHED, sheet_names)
        self.assertIn(SHEET_ALREADY_PAID, sheet_names)
        self.assertIn(SHEET_SUMMARY, sheet_names)

        # Check columns in master data sheet
        master_sheet_df = pd.read_excel(wb_file, sheet_name=SHEET_MASTER_DATA)
        for col in MASTER_COLUMNS:
            self.assertIn(col, master_sheet_df.columns)


if __name__ == "__main__":
    unittest.main()
