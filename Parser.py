from pathlib import Path
from collections import namedtuple

import openpyxl
from openpyxl.worksheet import Worksheet
from Application import Job
from Pre_processing import extract_month_and_year


class raw_data:
    def __init__(self):
        self.date = None
        self.object = None
        self.system = None
        self.work_type = None
        self.place = None


def xstr(cell_value):
    if cell_value is None:
        return None
    else:
        return str(cell_value)


def parser_asu(file_path: Path):
    sheet_names = {'АСУ ТП': 'АСУ ТП',
                   'АСУ И': 'АСУ И',
                   'МОСТ': 'АСУ АМ',
                   'ЛВС': 'ЛВС'}

    table_area = namedtuple('table_area', 'first_row first_col last_row last_col')

    def find_data_boundaries(sheet):
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
        data_area = table_area(first_row, first_col, last_row, last_col)
        return data_area()

    def parse_sheet(sheet: Worksheet, system: str) -> [Job]:
        # print(f'sheet: {sheet.title}')
        # print(f'system: {system}')

        data = find_data_boundaries(sheet)

        raw_month = sheet.cell(data_first_row - 2, data_first_col).value
        if raw_month is None:
            raw_month = sheet.cell(data_first_row - 3, data_first_col).value
        month, year = extract_month_and_year(raw_month)

        print(sheet.title, data_first_row, data_first_col, raw_month,
              month, year)  # debug

        print(f' end data col {last_col}; last date {sheet.cell(first_row - 1, last_col).value}')

        # for i_row in range(first_row, last_row + 1):
        #     place = sheet.cell(i_row, 2).value
        #     print(place)

    wb = openpyxl.load_workbook(str(file_path))
    for sheet_name in wb.sheetnames:  # find necessary worksheets by names
        for name, system in sheet_names.items():
            if name in sheet_name:
                parse_sheet(wb[sheet_name], system)
                break
