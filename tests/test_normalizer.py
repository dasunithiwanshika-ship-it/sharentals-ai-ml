"""Unit tests for normalization functions."""

import unittest
import pandas as pd
from toll_processor.normalizer import (
    normalize_rego,
    parse_flexible_datetime,
    normalize_amount,
    generate_duplicate_key,
)


class TestNormalizer(unittest.TestCase):
    def test_normalize_rego(self):
        self.assertEqual(normalize_rego("abc 123"), "ABC123")
        self.assertEqual(normalize_rego("  XYZ-789  "), "XYZ789")
        self.assertEqual(normalize_rego("1-abc.456"), "1ABC456")
        self.assertEqual(normalize_rego(None), "")
        self.assertEqual(normalize_rego(""), "")

    def test_parse_flexible_datetime(self):
        # Australian format DD/MM/YYYY
        dt1 = parse_flexible_datetime("08/09/2026 10:30:00")
        self.assertIsNotNone(dt1)
        self.assertEqual(dt1.day, 8)
        self.assertEqual(dt1.month, 9)
        self.assertEqual(dt1.year, 2026)
        self.assertEqual(dt1.hour, 10)
        self.assertEqual(dt1.minute, 30)

        # ISO format
        dt2 = parse_flexible_datetime("2026-09-08 14:00:00")
        self.assertIsNotNone(dt2)
        self.assertEqual(dt2.day, 8)
        self.assertEqual(dt2.month, 9)

        # Invalid / None
        self.assertIsNone(parse_flexible_datetime(None))
        self.assertIsNone(parse_flexible_datetime("not a date"))

    def test_normalize_amount(self):
        self.assertEqual(normalize_amount(-7.86, make_positive=True), 7.86)
        self.assertEqual(normalize_amount("-$7.86", make_positive=True), 7.86)
        self.assertEqual(normalize_amount("($5.55)", make_positive=True), 5.55)
        self.assertEqual(normalize_amount(10.504, make_positive=True), 10.50)
        self.assertEqual(normalize_amount(None), 0.0)

    def test_generate_duplicate_key(self):
        dt = pd.Timestamp("2026-09-08 10:30:00")
        key = generate_duplicate_key("ABC 123", dt, -7.86)
        self.assertEqual(key, "ABC123|2026-09-08 10:30:00|7.86")


if __name__ == "__main__":
    unittest.main()
