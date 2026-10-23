"""Regression checks using only temporary, fictional workbooks."""
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import openpyxl
from ranking_report import confirmed_count, generate_report


class ReportTests(unittest.TestCase):
    def setUp(self):
        self.folder = tempfile.TemporaryDirectory()
        self.path = Path(self.folder.name) / "test.xlsx"
        self.workbook = openpyxl.Workbook()
        self.sheet = self.workbook.active
        self.sheet.title = "October 2026"
        self.sheet.append(["Team A", None])
        self.sheet.append(["Agent", "Sales"])
        self.sheet.append(["Agent 01", 18])
        self.sheet["B3"].comment = openpyxl.comments.Comment("12 confirmed; confirmed", "Sample")
        self.sheet.append([None, None])
        self.sheet.append(["Agent 02", 0])
        self.sheet.append(["Agent 03", "-"])
        self.sheet.append(["TOTAL", "=SUM(B3:B6)"])

    def tearDown(self):
        self.workbook.close()
        self.folder.cleanup()

    def save_report(self, **kwargs):
        self.workbook.save(self.path)
        return generate_report(self.path, **kwargs)

    def test_multidigit_missing_and_repeated_notes(self):
        self.assertEqual(confirmed_count("12 confirmed; CONFIRMED; 3 confirmed"), 16)
        self.assertEqual(confirmed_count(None), 0)
        self.assertEqual(confirmed_count("follow-up needed; unconfirmed"), 0)
        self.assertEqual(confirmed_count("0 confirmed"), 0)

    def test_invalid_count_is_not_silently_misread(self):
        for note in ("-2 confirmed", "1.5 confirmed", "two confirmed", "12 confirmed today"):
            with self.subTest(note=note), self.assertRaises(ValueError):
                confirmed_count(note)

    def test_gaps_zero_sales_and_missing_comments(self):
        report = self.save_report()
        self.assertIn("Total sales: 18", report)
        self.assertIn("Total confirmed: 13", report)
        self.assertIn("Agent 02: 0 (No notes)", report)
        self.assertNotIn("Agent 03", report)

    def test_explicit_sheet_prevents_wrong_month(self):
        other = self.workbook.create_sheet("MARCH")
        other.append(["Team A"])
        other.append(["Agent", "Sales"])
        other.append(["Agent 01", 1])
        self.workbook.save(self.path)
        with self.assertRaisesRegex(ValueError, "Choose --sheet"):
            generate_report(self.path)
        self.assertIn("Total sales: 18", generate_report(self.path, "October 2026"))

    def test_missing_sheet_and_campaign_have_clear_errors(self):
        for kwargs, expected in (({"sheet_name": "Missing"}, "Sheet 'Missing' not found"),
                                 ({"campaigns": ["Missing"]}, "Campaign 'missing' not found")):
            with self.subTest(kwargs=kwargs), self.assertRaisesRegex(ValueError, expected):
                self.save_report(**kwargs)

    def test_invalid_sales(self):
        for value in (None, "", "abc", -1, 2.5, True, "NaN", "inf", "=1+1"):
            self.sheet["B3"] = value
            with self.subTest(value=value), self.assertRaisesRegex(ValueError, "B3"):
                self.save_report()

    def test_duplicate_campaign_or_agent_is_rejected(self):
        self.sheet["A5"] = "Agent 01"
        with self.assertRaisesRegex(ValueError, "duplicate agent"):
            self.save_report()
        self.sheet["A5"] = "Agent 02"
        self.sheet["D1"], self.sheet["D2"], self.sheet["E2"] = "team a", "Agent", "Sales"
        with self.assertRaisesRegex(ValueError, "Duplicate campaign"):
            self.save_report()

    def test_cli_keeps_existing_report_on_validation_failure(self):
        self.sheet["B3"] = "invalid"
        self.workbook.save(self.path)
        output = Path(self.folder.name) / "report.txt"
        output.write_text("existing report", encoding="utf-8")
        script = Path(__file__).resolve().parents[1] / "ranking_report.py"
        result = subprocess.run([sys.executable, str(script), str(self.path), "--output", str(output)],
                                capture_output=True, text=True)
        self.assertEqual(result.returncode, 1)
        self.assertIn("B3", result.stderr)
        self.assertEqual(output.read_text(encoding="utf-8"), "existing report")

    def test_committed_sample_matches_real_cli_output(self):
        root = Path(__file__).resolve().parents[1]
        output = Path(self.folder.name) / "sample.txt"
        result = subprocess.run([sys.executable, str(root / "ranking_report.py"),
                                 str(root / "samples/sample-ranking.xlsx"), "--sheet", "October 2026",
                                 "--output", str(output)], capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(output.read_text(encoding="utf-8"),
                         (root / "samples/sample-report.txt").read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
