from pathlib import Path
from datetime import date  # Посмотреть этот модуль

import openpyxl
from openpyxl.worksheet import Worksheet
import re


class Job:
    def __init__(self):
        self.date = None
        self.object = None
        self.system = None
        self.work_type = None
        self.place = None
        self.worker = None


def xstr(cell_value):
    if cell_value is None:
        return None
    else:
        return str(cell_value)


def extract_month_and_year(raw_date: str):
    month_names = {'январь': 1,
                   'февраль': 2,
                   'март': 3,
                   'апрель': 4,
                   'май': 5,
                   'июнь': 6,
                   'июль': 7,
                   'август': 8,
                   'сентябрь': 9,
                   'октябрь': 10,
                   'ноябрь': 11,
                   'декабрь': 12}
    for m_name, m_num in month_names.items():
        if m_name in raw_date.lower():
            month = int(m_num)
    year = int(re.findall(r'\d+', raw_date)[0])
    # print(raw_date, month, year)  # debug
    return month, year


def parser_asu(file_path: Path):
    sheet_names = {'АСУ ТП': 'АСУ ТП',
                   'АСУ И': 'АСУ И',
                   'МОСТ': 'АСУ АМ',
                   'ЛВС': 'ЛВС'}

    def parse_sheet(sheet: Worksheet, system: str) -> [Job]:
        # print(f'sheet: {sheet.title}')
        # print(f'system: {system}')
        for i_row in range(1, 30):
            if xstr(sheet.cell(i_row, 1).value) == "1":
                for i_col in range(1, 30):
                    if xstr(sheet.cell(i_row - 1, i_col).value) == "1":
                        origin = sheet.cell(i_row, i_col)
                        break
                break

        raw_month = sheet.cell(origin.row - 2, origin.col_idx).value
        if raw_month is None:
            raw_month = sheet.cell(origin.row - 3, origin.col_idx).value
        month, year = extract_month_and_year(raw_month)

        print(sheet.title, origin.row, origin.col_idx, raw_month,
              month, year)  # debug

        for i_row in range(origin.row, origin.row + 20):
            if sheet.cell(i_row,1) is None:
                end_data_row = i_row
                break
        print(end_data_row)

        for i_row in range(origin.row, origin.row + 40):
            place = sheet.cell(i_row, 2).value
            print(place)

    wb = openpyxl.load_workbook(str(file_path))
    # print(wb.sheetnames)  # debug
    for sheet_name in wb.sheetnames:  # find necessary worksheets by names
        for name, system in sheet_names.items():
            if name in sheet_name:
                parse_sheet(wb[sheet_name], system)


if __name__ == "__main__":
    jobs_schedule_asu = Path(r"c:\Users\Mihail\Documents\Гусев М.В\Планы работ\Графики ТО\5. Графики на 05.18 АСУ.XLSX")
    parser_asu(jobs_schedule_asu)
