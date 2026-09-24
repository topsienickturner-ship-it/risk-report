"""Run with: python -m unittest -v"""
import tempfile
import unittest
from pathlib import Path
from openpyxl import Workbook
from risk_report import Risk, read_risks, build_report


class ReportTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)

    def workbook(self, rows):
        wb = Workbook()
        ws = wb.active
        for _ in range(9):
            ws.append([])
        ws.append(['Risk ID', 'Title', 'Action ID', 'Action Title'])
        for row in rows:
            ws.append(row)
        path = self.root / 'input.xlsx'
        wb.save(path)
        return path

    def test_groups_non_adjacent_rows_and_deduplicates(self):
        path = self.workbook([['R1', 'One', 'A1', 'First'], ['R2', 'Two', None, None],
                              ['R1', 'One', 'A2', 'Second'], ['R1', 'One', 'A1', 'First']])
        risks = read_risks(path)
        self.assertEqual([r.risk_id for r in risks], ['R1', 'R2'])
        self.assertEqual(len(risks[0].actions), 2)
        self.assertEqual(risks[1].actions, [])

    def test_conflict(self):
        path = self.workbook([['R1', 'One', 'A1', 'First'], ['R1', 'Changed', 'A2', 'Second']])
        with self.assertRaisesRegex(ValueError, 'conflicting'):
            read_risks(path)

    def test_empty_and_missing_id(self):
        with self.assertRaisesRegex(ValueError, 'No risk records'):
            read_risks(self.workbook([]))
        with self.assertRaisesRegex(ValueError, 'no Risk ID'):
            read_risks(self.workbook([[None, 'One', 'A1', 'First']]))

    def test_conflicting_action(self):
        path = self.workbook([['R1', 'One', 'A1', 'First'], ['R1', 'One', 'A1', 'Changed']])
        with self.assertRaisesRegex(ValueError, 'conflicting duplicate action'):
            read_risks(path)

    def test_long_action_and_long_description_paginate(self):
        risks = [Risk('R1', {'Title': 'Long action', 'Description': 'A risk & its effects'},
                      [{'Action ID': 'A1', 'Action last comment': 'Long commentary. ' * 2000}]),
                 Risk('R2', {'Title': 'Long description', 'Description': 'Long description. ' * 2500}),
                 Risk('R3', {'Title': 'Short risk'})]
        output = self.root / 'report.pdf'
        pages = build_report(risks, output)
        self.assertGreater(pages.count('R1'), 1)
        self.assertGreater(pages.count('R2'), 1)
        self.assertEqual(pages[-1], 'R3')
        self.assertEqual(pages, sorted(pages))
        self.assertEqual(pages.count('R3'), 1)
        self.assertTrue(output.read_bytes().startswith(b'%PDF-'))


if __name__ == '__main__':
    unittest.main()
