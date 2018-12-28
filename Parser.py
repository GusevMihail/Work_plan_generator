from pathlib import Path

import openpyxl
from openpyxl.worksheet import Worksheet
from Application import Job, extract_month_and_year

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

    def parse_sheet(sheet: Worksheet, system: str) -> [Job]:
        # print(f'sheet: {sheet.title}')
        # print(f'system: {system}')
        def find_data_boundaries():
        data_first_row = 1
        data_first_col = 1
        for i_row in range(1, 30):
            if xstr(sheet.cell(i_row, 1).value) == "1":
                for i_col in range(1, 30):
                    if xstr(sheet.cell(i_row - 1, i_col).value) == "1":
                        data_first_row = i_row
                        data_first_col = i_col
                        break
                break

        data_last_row = data_first_row
        for i_row in range(data_first_row, data_first_row + 20):
            if sheet.cell(i_row, 1).value is None:
                data_last_row = i_row - 1
                break

        data_last_col = data_first_col
        for i_col in range(data_first_col, data_first_col + 40):
            if sheet.cell(data_first_row - 1, i_col).value is None:
                data_last_col = i_col - 1
                break

        raw_month = sheet.cell(data_first_row - 2, data_first_col).value
        if raw_month is None:
            raw_month = sheet.cell(data_first_row - 3, data_first_col).value
        month, year = extract_month_and_year(raw_month)

        print(sheet.title, data_first_row, data_first_col, raw_month,
              month, year)  # debug

        print(f' end data col {data_last_col}; last date {sheet.cell(data_first_row - 1, data_last_col).value}')

        # for i_row in range(data_first_row, data_last_row + 1):
        #     place = sheet.cell(i_row, 2).value
        #     print(place)

    wb = openpyxl.load_workbook(str(file_path))
    for sheet_name in wb.sheetnames:  # find necessary worksheets by names
        for name, system in sheet_names.items():
            if name in sheet_name:
                parse_sheet(wb[sheet_name], system)
                break
