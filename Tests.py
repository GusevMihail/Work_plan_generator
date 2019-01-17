import unittest
from pathlib import Path
import Parser
import openpyxl


class TestParser(unittest.TestCase):
    wb_asu = openpyxl.load_workbook(r'.\input data\5. Графики на 05.18 АСУ.xlsx')
    wb_test = openpyxl.load_workbook(r'.\input data\Test Schedule.xlsx')
    sheet_asu_109 = wb_asu['109. МОСТ']
    sheet_asu_107 = wb_asu['107. АСУ ТП']
    sheet_test_107 = wb_test['107. АСУ ТП']
    parser_asu_109 = Parser.ParserAsu(sheet_asu_109)
    parser_asu_107 = Parser.ParserAsu(sheet_asu_107)
    parser_test_107 = Parser.ParserAsu(sheet_test_107)

    def test_find_data_boundaries(self):
        self.parser_asu_109.find_data_boundaries()
        self.assertEqual(self.parser_asu_109.data_area, (8, 6, 8, 36))
        self.parser_asu_107.find_data_boundaries()
        self.assertEqual(self.parser_asu_107.data_area, (5, 3, 15, 33))

    def test_find_month_year(self):
        self.parser_asu_107.find_data_boundaries()
        self.parser_asu_107.find_month_year()
        self.assertEqual(self.parser_asu_107.month_year, 'МАЙ  2018 г.')
        self.parser_asu_109.find_data_boundaries()
        self.parser_asu_109.find_month_year()
        self.assertEqual(self.parser_asu_109.month_year, 'май 2018г.')

    def test_extract_jobs(self):
        self.parser_test_107.find_data_boundaries()
        self.parser_test_107.extract_jobs()
        self.assertEqual(self.parser_test_107.raw_data, ())


if __name__ == '__main__':
    unittest.main()
