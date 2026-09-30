"""Unit tests for duplicate and already-paid detector."""

import unittest
import pandas as pd
from toll_processor.duplicate_detector import DuplicateDetector
from toll_processor.normalizer import generate_duplicate_key
from toll_processor.schemas import STATUS_DUPLICATE, STATUS_ALREADY_PAID


class TestDuplicateDetector(unittest.TestCase):
    def setUp(self):
        # Historical master
        self.master_df = pd.DataFrame([
            {
                "REGO": "HIST001",
                "StartDateTime": "22/08/2026 10:00:00",
                "Toll Amount": 5.12,
                "Payment Status": "Pending",
                "Payment Method": "",
            },
            {
                "REGO": "HIST001",
                "StartDateTime": "23/08/2026 11:30:00",
                "Toll Amount": 7.86,
                "Payment Status": "Paid",
                "Payment Method": "Card",
            },
        ])
        self.detector = DuplicateDetector(self.master_df)

    def test_historical_duplicate_unpaid(self):
        key = generate_duplicate_key("HIST001", pd.Timestamp("2026-08-22 10:00:00"), 5.12)
        is_dup, status, reason = self.detector.check_duplicate(key)
        self.assertTrue(is_dup)
        self.assertEqual(status, STATUS_DUPLICATE)

    def test_historical_already_paid(self):
        key = generate_duplicate_key("HIST001", pd.Timestamp("2026-08-23 11:30:00"), 7.86)
        is_dup, status, reason = self.detector.check_duplicate(key)
        self.assertTrue(is_dup)
        self.assertEqual(status, STATUS_ALREADY_PAID)

    def test_new_toll_and_batch_duplicate(self):
        key = generate_duplicate_key("NEW001", pd.Timestamp("2026-09-08 10:00:00"), 12.50)
        # First time: Not duplicate
        is_dup1, status1, _ = self.detector.check_duplicate(key)
        self.assertFalse(is_dup1)
        self.assertIsNone(status1)

        # Second time in same batch: Duplicate
        is_dup2, status2, reason2 = self.detector.check_duplicate(key)
        self.assertTrue(is_dup2)
        self.assertEqual(status2, STATUS_DUPLICATE)
        self.assertIn("batch", reason2.lower())


if __name__ == "__main__":
    unittest.main()
