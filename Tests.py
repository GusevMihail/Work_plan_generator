import unittest
# from pathlib import Path
import openpyxl
import Parser
import Pre_processing


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
        self.parser_asu_107.find_month_year()
        self.assertEqual(self.parser_asu_107.month_year, 'МАЙ  2018 г.')
        self.parser_asu_109.find_month_year()
        self.assertEqual(self.parser_asu_109.month_year, 'май 2018г.')

    def test_extract_jobs(self):
        self.parser_test_107.extract_jobs()
        last = len(self.parser_test_107.raw_data) - 1
        self.assertEqual(self.parser_test_107.raw_data[0].day, 1)
        self.assertEqual(self.parser_test_107.raw_data[0].work_type, 'ТО1')
        self.assertEqual(self.parser_test_107.raw_data[0].place, 'ПТК судопропускного сооружения - ПТК С1 север')
        self.assertEqual(self.parser_test_107.raw_data[last].day, 31)
        self.assertEqual(self.parser_test_107.raw_data[last].work_type, 'ТО4')
        self.assertEqual(self.parser_test_107.raw_data[last].place, 'ПТК ЗУ КЗС')


class TestPreProcessing(unittest.TestCase):
#     test_raw_places = open(r'.\input data\test raw places.txt')
#     for line in test_raw_places:
#         print(f'{line}  -->>  { extract_place(line)}')

    def test_extract_system(self):
        self.assertEqual(Pre_processing.extract_system('107. АСУ ТП'), 'АСУ ТП')
        self.assertEqual(Pre_processing.extract_system('108. АСУ И '), 'АСУ И')
        self.assertEqual(Pre_processing.extract_system('109. МОСТ'), 'АСУ АМ')
        self.assertEqual(Pre_processing.extract_system('Лист 1'), None)
        self.assertEqual(Pre_processing.extract_system(''), None)


if __name__ == '__main__':
    unittest.main()

    # TP = TestParser()
    # TP.parser_test_107.find_data_boundaries()
    # TP.parser_test_107.extract_jobs()
    # # print(TP.parser_test_107.raw_data)
    # print(len(TP.parser_test_107.raw_data))
