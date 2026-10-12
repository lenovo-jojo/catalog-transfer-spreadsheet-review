"""End-to-end checks using entirely invented catalog rows."""
import csv
import os
from pathlib import Path
import subprocess
import tempfile
import unittest
import zipfile
from xml.sax.saxutils import escape


def workbook(path, rows):
    xml_rows = []
    for index, row in enumerate(rows, 1):
        cells = []
        for column, value in row.items():
            cells.append(f'<c r="{column}{index}" t="inlineStr"><is><t>{escape(value)}</t></is></c>')
        xml_rows.append(f'<row r="{index}">{"".join(cells)}</row>')
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("xl/workbook.xml", '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Report" sheetId="1" r:id="rId2"/></sheets></workbook>')
        z.writestr("xl/_rels/workbook.xml.rels", '<Relationships><Relationship Id="rId2" Target="worksheets/report.xml"/></Relationships>')
        z.writestr("xl/worksheets/report.xml", f'<worksheet><sheetData>{"".join(xml_rows)}</sheetData></worksheet>')


class ClaimantsTests(unittest.TestCase):
    def run_tool(self, path):
        return subprocess.run([os.environ["CLAIMANTS_BINARY"], str(path)], env={**os.environ, "CLAIMANTS_NO_CLIPBOARD": "1"}, text=True, capture_output=True)

    def test_filter_deduplicate_and_csv_escape(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "synthetic.xlsx"
            header = {"E": "Transfer", "K": "Member ID", "L": "Member", "M": "Publisher ID", "N": "Publisher", "Q": "Admin ID", "R": "Admin"}
            claimant = {"K": "001", "L": 'Example, "Studio"', "M": "002", "N": "Line\nTwo", "Q": "003", "R": "Synthetic Admin"}
            different = {**claimant, "K": "004"}
            workbook(path, [header, {**claimant, "E": " Yes "}, {**claimant, "E": "(Y)"}, {**different, "E": "y"}, {**different, "E": "No"}])
            result = self.run_tool(path)
            self.assertEqual(result.returncode, 0, result.stderr)
            with path.with_name("synthetic_yes_claimants.csv").open(newline="") as f:
                rows = list(csv.reader(f))
            self.assertEqual(len(rows), 3)
            self.assertEqual(rows[1], [claimant[c] for c in ["K", "L", "M", "N", "Q", "R"]])
            self.assertEqual(rows[2][0], "004")

    def test_no_matches_reports_error(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "synthetic.xlsx"
            workbook(path, [{"E": "Transfer"}, {"E": "No"}])
            result = self.run_tool(path)
            self.assertNotEqual(result.returncode, 0)
            self.assertIn("No YES", result.stderr)

    def test_rejects_other_file_types(self):
        self.assertNotEqual(self.run_tool(Path("input.csv")).returncode, 0)


if __name__ == "__main__":
    unittest.main()
