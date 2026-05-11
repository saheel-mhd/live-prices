"""Tiny smoke tests for liveprices: price formatting and config round-trip.

Run with:  python -m unittest test_liveprices    (no extra deps needed)
Windows-only — importing liveprices loads the Win32 DDEML bindings.
"""
import os
import tempfile
import unittest

import liveprices


class FmtTests(unittest.TestCase):
    def test_blank_and_none(self):
        self.assertEqual(liveprices._fmt(None), "")
        self.assertEqual(liveprices._fmt(""), "")

    def test_non_numeric_passthrough(self):
        self.assertEqual(liveprices._fmt("N/A"), "N/A")

    def test_decimals_by_magnitude(self):
        self.assertEqual(liveprices._fmt(1.2345678), "1.234568")    # < 99   -> 6 dp
        self.assertEqual(liveprices._fmt(150.123456), "150.12346")  # 99..998 -> 5 dp
        self.assertEqual(liveprices._fmt(1234.5), "1234.5000")      # 999..9999 -> 4 dp
        self.assertEqual(liveprices._fmt(12345.6), "12345.600")     # > 9999  -> 3 dp

    def test_accepts_string_numbers(self):
        self.assertEqual(liveprices._fmt("1.23456789"), "1.234568")


class ConfigRoundTripTests(unittest.TestCase):
    def setUp(self):
        self._orig = liveprices.CONFIG_FILE
        fd, self.path = tempfile.mkstemp(suffix=".txt")
        os.close(fd)
        liveprices.CONFIG_FILE = self.path

    def tearDown(self):
        liveprices.CONFIG_FILE = self._orig
        try:
            os.remove(self.path)
        except OSError:
            pass

    def test_save_then_load(self):
        pairs = [("GOLD", "XAUUSD"), ("EUR", "EURUSD")]
        liveprices.save_config(pairs=pairs, rows=["GOLD", "EUR", ""], is_darkmode=False)
        cfg = liveprices.load_config()
        self.assertEqual(cfg["PAIRS"], pairs)
        self.assertEqual(cfg["ROWS"], ["GOLD", "EUR", ""])
        self.assertIs(cfg["IS_DARKMODE"], False)

    def test_legacy_single_value_symbol(self):
        with open(self.path, "w", encoding="utf-8") as f:
            f.write("SYMBOLS=EURUSD\n")
        cfg = liveprices.load_config()
        self.assertEqual(cfg["PAIRS"], [("EURUSD", "EURUSD")])

    def test_missing_file_returns_none(self):
        os.remove(self.path)
        self.assertIsNone(liveprices.load_config())


if __name__ == "__main__":
    unittest.main()
