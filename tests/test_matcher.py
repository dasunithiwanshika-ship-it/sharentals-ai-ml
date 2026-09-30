"""Unit tests for driver matching engine."""

import unittest
import pandas as pd
from toll_processor.matcher import match_driver_for_toll
from toll_processor.schemas import STATUS_UNMATCHED, STATUS_MULTIPLE_MATCH


class TestMatcher(unittest.TestCase):
    def setUp(self):
        # Sample bookings
        self.bookings_df = pd.DataFrame([
            {
                "norm_rego": "ABC123",
                "client": "John Smith",
                "booking_ref": "BK-001",
                "start_dt": pd.Timestamp("2026-09-07 09:00:00"),
                "finish_dt": pd.Timestamp("2026-09-10 17:00:00"),
                "contact": "0412345678",
            },
            {
                "norm_rego": "OVERLAP1",
                "client": "Driver A",
                "booking_ref": "BK-002A",
                "start_dt": pd.Timestamp("2026-09-08 08:00:00"),
                "finish_dt": pd.Timestamp("2026-09-08 18:00:00"),
                "contact": "0400000001",
            },
            {
                "norm_rego": "OVERLAP1",
                "client": "Driver B",
                "booking_ref": "BK-002B",
                "start_dt": pd.Timestamp("2026-09-08 12:00:00"),
                "finish_dt": pd.Timestamp("2026-09-09 12:00:00"),
                "contact": "0400000002",
            },
        ])

    def test_single_match(self):
        toll = pd.Series({
            "raw_rego": "ABC 123",
            "norm_rego": "ABC123",
            "parsed_start_dt": pd.Timestamp("2026-09-08 10:30:00"),
            "toll_amount": 7.86,
        })
        res = match_driver_for_toll(toll, self.bookings_df)
        self.assertTrue(res["matched"])
        self.assertEqual(res["driver"], "John Smith")
        self.assertEqual(res["booking_ref"], "BK-001")
        self.assertEqual(res["hire_date"], pd.Timestamp("2026-09-07 09:00:00"))
        self.assertEqual(res["return_date"], pd.Timestamp("2026-09-10 17:00:00"))

    def test_out_of_bounds_no_match(self):
        # Toll happened before start date
        toll = pd.Series({
            "raw_rego": "ABC 123",
            "norm_rego": "ABC123",
            "parsed_start_dt": pd.Timestamp("2026-09-05 10:30:00"),
            "toll_amount": 7.86,
        })
        res = match_driver_for_toll(toll, self.bookings_df)
        self.assertFalse(res["matched"])
        self.assertEqual(res["driver"], "UNMATCHED")
        self.assertEqual(res["status"], STATUS_UNMATCHED)

    def test_unknown_rego(self):
        toll = pd.Series({
            "raw_rego": "UNKNOWN99",
            "norm_rego": "UNKNOWN99",
            "parsed_start_dt": pd.Timestamp("2026-09-08 10:30:00"),
            "toll_amount": 7.86,
        })
        res = match_driver_for_toll(toll, self.bookings_df)
        self.assertFalse(res["matched"])
        self.assertIn("No matching Rego", res["reason"])

    def test_multiple_match(self):
        # At 14:00 on 08/09/2026, both Driver A and Driver B bookings are active
        toll = pd.Series({
            "raw_rego": "OVERLAP1",
            "norm_rego": "OVERLAP1",
            "parsed_start_dt": pd.Timestamp("2026-09-08 14:00:00"),
            "toll_amount": 8.50,
        })
        res = match_driver_for_toll(toll, self.bookings_df)
        self.assertFalse(res["matched"])
        self.assertEqual(res["status"], STATUS_MULTIPLE_MATCH)
        self.assertEqual(res["driver"], "UNMATCHED")
        self.assertIn("Multiple active bookings", res["reason"])


if __name__ == "__main__":
    unittest.main()
