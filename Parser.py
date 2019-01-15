from pathlib import Path
from collections import namedtuple

import openpyxl
from openpyxl.worksheet import Worksheet
from Application import Job
from Pre_processing import extract_month_and_year


class RawData:
    def __init__(self):
        self.date = None
        self.object = None
        self.system = None
        self.work_type = None
        self.place = None
        self.sheet_name = None



def xstr(cell_value):
    if cell_value is None:
        return None
    else:
        return str(cell_value)


table_area = namedtuple('table_area', 'first_row first_col last_row last_col')


class ParserAsu:
    sheet_names = {'АСУ ТП': 'АСУ ТП',
                   'АСУ И': 'АСУ И',
                   'МОСТ': 'АСУ АМ',
                   'ЛВС': 'ЛВС'}

    def __init__(self):
        self.sheet = None
        self.data_area = None
        self.raw_data = []

    def find_data_boundaries(self, sheet):
        first_row = 1
        first_col = 1
        for i_row in range(1, 30):
            if xstr(sheet.cell(i_row, 1).value) == "1":
                for i_col in range(1, 30):
                    if xstr(sheet.cell(i_row - 1, i_col).value) == "1":
                        first_row = i_row
                        first_col = i_col
                        break
                break
        last_row = first_row
        for i_row in range(first_row, first_row + 20):
            if sheet.cell(i_row, 1).value is None:
                last_row = i_row - 1
                break
        last_col = first_col
        for i_col in range(first_col, first_col + 40):
            if sheet.cell(first_row - 1, i_col).value is None:
                last_col = i_col - 1
                break
        self.data_area = table_area(first_row, first_col, last_row, last_col)

    def parse_sheet(self, sheet: Worksheet, system: str):
        # print(f'sheet: {sheet.title}')
        # print(f'system: {system}')

        raw_month = sheet.cell(self.data_area.first_row - 2, self.data_area.first_col).value
        if raw_month is None:
            raw_month = sheet.cell(self.data_area.first_row - 3, self.data_area.first_col).value
        month, year = extract_month_and_year(raw_month)

        # print(sheet.title, data_first_row, data_first_col, raw_month,
        #       month, year)  # debug
        #
        # print(f' end data col {last_col}; last date {sheet.cell(first_row - 1, last_col).value}')

        # for i_row in range(first_row, last_row + 1):
        #     place = sheet.cell(i_row, 2).value
        #     print(place)


