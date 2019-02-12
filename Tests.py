import datetime
import unittest
from typing import List

import openpyxl
from openpyxl.styles import NamedStyle

import Application
import Parser
import Pre_processing
import Cell_styler
import Table_generator


class TestParserAsu(unittest.TestCase):
    wb_asu = openpyxl.load_workbook(r'.\input data\5. Графики на 05.18 АСУ.xlsx')
    wb_test = openpyxl.load_workbook(r'.\input data\Test Schedule.xlsx')
    sheet_asu_109 = wb_asu['109. МОСТ']
    sheet_asu_107 = wb_asu['107. АСУ ТП']
    sheet_test_107 = wb_test['107. АСУ ТП']
    parser_asu_109 = Parser.ParserAsu(sheet_asu_109)
    parser_asu_107 = Parser.ParserAsu(sheet_asu_107)
    parser_test_107 = Parser.ParserAsu(sheet_test_107)

    def test_find_data_boundaries(self):
        self.assertEqual(self.parser_asu_109._data_area, (8, 6, 8, 36))
        self.assertEqual(self.parser_asu_107._data_area, (5, 3, 15, 33))

    def test_find_month_year(self):
        self.assertEqual(self.parser_asu_107.month_year, 'МАЙ  2018 г.')
        self.assertEqual(self.parser_asu_109.month_year, 'май 2018г.')

    def test_extract_jobs(self):
        last = len(self.parser_test_107.raw_data) - 1
        self.assertEqual(self.parser_test_107.raw_data[0].day, 1)
        self.assertEqual(self.parser_test_107.raw_data[0].work_type, 'ТО1')
        self.assertEqual(self.parser_test_107.raw_data[0].place, 'ПТК судопропускного сооружения - ПТК С1 север')
        self.assertEqual(self.parser_test_107.raw_data[last].day, 31)
        self.assertEqual(self.parser_test_107.raw_data[last].work_type, 'ТО4')
        self.assertEqual(self.parser_test_107.raw_data[last].place, 'ПТК ЗУ КЗС')


class TestParserVols(unittest.TestCase):
    wb_vols = openpyxl.load_workbook(r'.\input data\Test Schedule VOLS.xlsx')
    sheet_1 = wb_vols['8.1.38 ТО']
    sheet_2 = wb_vols['10.4.38 ТО']
    parser_1 = Parser.ParserVOLS(sheet_1)
    parser_2 = Parser.ParserVOLS(sheet_2)

    def test_find_data_boundaries(self):
        self.assertEqual(self.parser_1._data_first_col, 7)
        self.assertEqual(self.parser_1._data_last_col, 37)
        self.assertEqual(self.parser_1._data_rows, [20, 21])
        self.assertEqual(self.parser_2._data_first_col, 7)
        self.assertEqual(self.parser_2._data_last_col, 37)
        self.assertEqual(self.parser_2._data_rows, [22, 23, 26])

    def test_find_month_year(self):
        self.assertEqual(self.parser_1.month_year, 'Май 2018 года')
        self.assertEqual(self.parser_2.month_year, 'Май 2018 года')

    def test_extract_jobs(self):
        self.assertEqual(self.parser_1.raw_data[0].day, 11)
        self.assertEqual(self.parser_1.raw_data[0].work_type, 'ТО2')
        self.assertEqual(self.parser_1.raw_data[0].place,
                         'Местоположение: Здание управления комплекса защитных сооружений')
        self.assertEqual(self.parser_1.raw_data[1].day, 18)
        self.assertEqual(self.parser_1.raw_data[1].work_type, 'ТО3')
        self.assertEqual(self.parser_1.raw_data[1].place,
                         'Местоположение: Здание управления комплекса защитных сооружений')

        self.assertEqual(self.parser_2.raw_data[0].day, 17)
        self.assertEqual(self.parser_2.raw_data[0].work_type, 'ТО2')
        self.assertEqual(self.parser_2.raw_data[0].place,
                         'Местоположение: Здание трансформаторной подстанции 110/10кВ '
                         'ПС С1 судопропускного сооружения С-1')
        self.assertEqual(self.parser_2.raw_data[2].day, 24)
        self.assertEqual(self.parser_2.raw_data[2].work_type, 'ТО2')
        self.assertEqual(self.parser_2.raw_data[2].place, 'ПС №86')


class TestPreProcessingAsu(unittest.TestCase):
    #     test_raw_places = open(r'.\input data\test raw places.txt')
    #     for line in test_raw_places:
    #         print(f'{line}  -->>  { extract_place_and_object(line)}')
    wb_asu = openpyxl.load_workbook(r'.\input data\5. Графики на 05.18 АСУ.xlsx')
    wb_test = openpyxl.load_workbook(r'.\input data\Test Schedule.xlsx')
    sheet_asu_109 = wb_asu['109. МОСТ']
    sheet_asu_107 = wb_asu['107. АСУ ТП']
    sheet_test_107 = wb_test['107. АСУ ТП']
    parser_asu_109 = Parser.ParserAsu(sheet_asu_109)
    parser_asu_107 = Parser.ParserAsu(sheet_asu_107)
    parser_test_107 = Parser.ParserAsu(sheet_test_107)

    def test_extract_system(self):
        self.assertEqual(Pre_processing.extract_system('107. АСУ ТП'), 'АСУ ТП')
        self.assertEqual(Pre_processing.extract_system('108. АСУ И '), 'АСУ И')
        self.assertEqual(Pre_processing.extract_system('109. МОСТ'), 'АСУ АМ')
        self.assertEqual(Pre_processing.extract_system('Лист 1'), None)
        self.assertEqual(Pre_processing.extract_system(''), None)

    def test_parser_to_jobs(self):
        jobs: List[Pre_processing.Job] = []
        jobs.extend(Pre_processing.parser_to_jobs(self.parser_test_107))
        last = len(jobs) - 1
        self.assertEqual(jobs[0].object, 'Судопропускное сооружение С1')
        self.assertEqual(jobs[0].place, 'С1 Север')
        self.assertEqual(jobs[0].work_type, 'ТО1')
        self.assertEqual(jobs[0].date, datetime.date(2018, 5, 1))
        self.assertEqual(jobs[0].system, 'АСУ ТП')
        self.assertEqual(jobs[last].object, 'Здание управления КЗС')
        self.assertEqual(jobs[last].place, 'Здание управления КЗС')
        self.assertEqual(jobs[last].work_type, 'ТО4')
        self.assertEqual(jobs[last].date, datetime.date(2018, 5, 31))
        self.assertEqual(jobs[last].system, 'АСУ ТП')

    def test_filter_work_type(self):
        self.assertEqual(Pre_processing.filter_work_type('ТО1'), 'ТО1')  # Rus to Rus
        self.assertEqual(Pre_processing.filter_work_type('TO2'), 'ТО2')  # Eng to Rus
        self.assertEqual(Pre_processing.filter_work_type('ETO'), 'ЕТО')  # Eng to Rus
        self.assertEqual(Pre_processing.filter_work_type('ЕТО \nТО1'), 'ЕТО \nТО1')  # Eng+Rus to Rus


class TestTableGenerator(unittest.TestCase):
    style1 = NamedStyle(name='style1')

    def test_apply_style(self):
        test_out_wb = openpyxl.Workbook()
        ws = test_out_wb.active
        area = Cell_styler.TableArea(first_row=3, last_row=6, first_col=2, last_col=4)
        Cell_styler.apply_named_style(ws, self.style1, table_area=area)
        for row in range(area.first_row, area.last_row + 1):
            for col in range(area.first_col, area.last_col + 1):
                self.assertEqual(ws.cell(row, col).style, self.style1.name)


class TestApplicationFunctions(unittest.TestCase):
    wb_test_asu = openpyxl.load_workbook(r'.\input data\Test Schedule.xlsx')
    wb_test_vols = openpyxl.load_workbook(r'.\input data\Test Schedule VOLS.xlsx')

    def test_find_sheets_asu(self):
        sheets = Application.find_sheets_asu(self.wb_test_asu)
        self.assertEqual(sheets[0].title, '107. АСУ ТП')
        self.assertEqual(sheets[1].title, '108. АСУ И ')
        self.assertEqual(sheets[2].title, '109. МОСТ')
        self.assertEqual(sheets[3].title, '110. ЛВС ')
        self.assertEqual(len(sheets), 4)

    def test_find_sheets_vols(self):
        sheets = Application.find_sheets_vols(self.wb_test_vols)
        self.assertEqual(sheets[0].title, '8.1.38 ТО')
        self.assertEqual(sheets[1].title, '10.2.38 ТО')
        self.assertEqual(sheets[2].title, '10.3.38 ТО')
        self.assertEqual(sheets[3].title, '10.4.38 ТО')
        self.assertEqual(len(sheets), 4)


if __name__ == '__main__':
    unittest.main()

    # TP = TestParser()
    # jobs = []
    # Pre_processing.parser_to_jobs(TP.parser_test_107, jobs)
