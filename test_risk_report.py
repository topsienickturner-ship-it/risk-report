"""Run with: python -m unittest -v"""
import tempfile
import unittest
import zipfile
import re
from pathlib import Path
from openpyxl import Workbook
from risk_report import Risk, read_risks, build_report, split_actions


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

    def test_blank_column_a_and_detect_header(self):
        wb = Workbook()
        ws = wb.active
        ws.append([None, 'Export title'])
        ws.append([None, 'Risk ID', 'Title', 'Action Title'])
        ws.append([None, 'R1', 'Example', '01 First\n02 Second'])
        path = self.root / 'leading-blank.xlsx'
        wb.save(path)
        risks = read_risks(path)
        self.assertEqual(risks[0].risk_id, 'R1')
        self.assertEqual([a['Action Title'] for a in risks[0].actions], ['First', 'Second'])
        self.assertEqual(len(read_risks(path, header_row=2)), 1)

    def test_incorrect_export_dimensions(self):
        path = self.workbook([['R1', 'One', 'A1', 'First']])
        with zipfile.ZipFile(path) as source:
            files = {name: source.read(name) for name in source.namelist()}
        name = 'xl/worksheets/sheet1.xml'
        files[name] = re.sub(rb'<dimension ref="[^"]+"', b'<dimension ref="A1"', files[name])
        with zipfile.ZipFile(path, 'w') as target:
            for name, content in files.items():
                target.writestr(name, content)
        self.assertEqual(read_risks(path)[0].risk_id, 'R1')

    def test_numbered_columns_align_by_number(self):
        actions = split_actions({'Action Title': '01. First\ncontinued text\n02. Second\n03. Third',
                                 'Action Owner': '03 Carol\n01 Alice',
                                 'Action Start Date (forecast)': '01 01/03/2026\n03 02/04/2026'})
        self.assertEqual(len(actions), 3)
        self.assertEqual(actions[0]['Action Title'], 'First\ncontinued text')
        self.assertEqual(actions[0]['Action Owner'], 'Alice')
        self.assertEqual(actions[1]['Action Owner'], '')
        self.assertEqual(actions[2]['Action Owner'], 'Carol')
        self.assertEqual(actions[2]['Action Start Date (forecast)'], '02/04/2026')

    def test_unnumbered_dates_and_multiline_text_preserved(self):
        record = {'Action Title': 'Review\nthe schedule',
                  'Action Start Date (forecast)': '01/03/2026'}
        self.assertEqual(split_actions(record), [record])

    def test_ambiguous_numbered_actions_rejected(self):
        with self.assertRaisesRegex(ValueError, 'unnumbered value'):
            split_actions({'Action Title': '01 First\n02 Second', 'Action Owner': 'Alice'})
        with self.assertRaisesRegex(ValueError, 'repeated action prefix'):
            split_actions({'Action Title': '01 First\n01 Second'})

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
