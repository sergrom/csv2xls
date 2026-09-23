"""End-to-end checks using an independent BIFF8 reader."""
import csv
import io
from pathlib import Path
import subprocess
import tempfile
import unittest

import xlrd

ROOT = Path(__file__).resolve().parents[1]


class ConversionTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.build_dir = tempfile.TemporaryDirectory()
        cls.addClassCleanup(cls.build_dir.cleanup)
        cls.binary = Path(cls.build_dir.name) / "csv2xls"
        subprocess.run(["go", "build", "-o", str(cls.binary), "."], cwd=ROOT, check=True)

    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.source = Path(self.directory.name) / "input.csv"
        self.output = Path(self.directory.name) / "output.xls"

    def convert(self, content, *args):
        self.source.write_bytes(content.encode("utf-8"))
        subprocess.run([
            str(self.binary), "--csv-file-name", str(self.source),
            "--xls-file-name", str(self.output), *args,
        ], check=True, capture_output=True, text=True)
        book = xlrd.open_workbook(self.output)
        self.addCleanup(book.release_resources)
        return book

    def test_values_and_bom(self):
        for prefix in ("", "\ufeff"):
            for delimiter in (";", ",", "§"):
                with self.subTest(prefix=repr(prefix), delimiter=delimiter):
                    rows = [
                        ["City", "Code", "Note"],
                        ["Paris", "001", "=1+1"],
                        ["Berlin", "🐧", 'a"quote'],
                        ["Madrid", "line1\nline2", "embedded\ufeffBOM"],
                        ["Short"],
                        ["Long", "é" * 32767, "🐧" * 16000],
                    ]
                    text = io.StringIO()
                    csv.writer(text, delimiter=delimiter, quoting=csv.QUOTE_ALL).writerows(rows)
                    book = self.convert(prefix + text.getvalue(), "--csv-delimiter", delimiter)
                    self.assertEqual(book.sheet_names(), ["worksheet"])
                    sheet = book.sheet_by_index(0)
                    self.assertEqual((sheet.nrows, sheet.ncols), (len(rows), 3))
                    for i, row in enumerate(rows):
                        self.assertEqual(sheet.row_values(i), row + [""] * (3 - len(row)))
                        for j in range(len(row)):
                            self.assertEqual(sheet.cell_type(i, j), xlrd.XL_CELL_TEXT)

    def test_sheet_boundaries(self):
        for count in (65535, 65536):
            with self.subTest(count=count):
                book = self.convert("".join(f"row {i}\n" for i in range(count)))
                expected_names = ["worksheet"] + (["worksheet1"] if count > 65535 else [])
                self.assertEqual(book.sheet_names(), expected_names)
                actual = []
                for sheet in book.sheets():
                    self.assertEqual(sheet.ncols, 1)
                    actual.extend(sheet.col_values(0))
                self.assertEqual(actual, [f"row {i}" for i in range(count)])
                self.assertEqual(book.sheet_by_index(0).nrows, 65535)
                if count > 65535:
                    self.assertEqual(book.sheet_by_index(1).nrows, 1)

    def test_empty_and_short_input(self):
        for content in ("", "\ufeff", "x", "é"):
            with self.subTest(content=repr(content)):
                book = self.convert(content)
                self.assertEqual(book.sheet_names(), ["worksheet"])
                sheet = book.sheet_by_index(0)
                self.assertEqual(sheet.nrows, 0 if content in ("", "\ufeff") else 1)
                if sheet.nrows:
                    self.assertEqual(sheet.cell_value(0, 0), content)

    def test_column_limit(self):
        values = [f"column {i}" for i in range(256)]
        book = self.convert(";".join(values))
        self.assertEqual(book.sheet_by_index(0).row_values(0), values)

    def test_cli_errors(self):
        self.source.write_text("City;Code\nParis;001\n", encoding="utf-8")
        for extra in (["unexpected"], ["--csv-delimiter=ab"]):
            with self.subTest(extra=extra):
                result = subprocess.run([
                    str(self.binary), "--csv-file-name", str(self.source),
                    "--xls-file-name", str(self.output), *extra,
                ], capture_output=True, text=True)
                self.assertEqual(result.returncode, 1)
                self.assertEqual(result.stdout, "")
                self.assertEqual(result.stderr.count("Error:"), 1)
                self.assertNotIn("Usage:", result.stderr)
                self.assertFalse(self.output.exists())
