from pathlib import Path
from datetime import date # Посмотреть этот модуль

import openpyxl
from openpyxl.worksheet import Worksheet


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


def parser_asu(file_path: Path):
    sheet_names = {'АСУ ТП': 'АСУ ТП',
                   'АСУ И': 'АСУ И',
                   'МОСТ': 'АСУ АМ',
                   'ЛВС': 'ЛВС'}

    def parse_sheet(sheet: Worksheet, system: str) -> [Job]:
        # print(f'sheet: {sheet.title}')
        # print(f'system: {system}')
        first_cell = None
        for i in range(1, 30):
            if xstr(sheet.cell(i, 1).value) == "1":
                for j in range(1, 30):
                    if xstr(sheet.cell(i - 1, j).value) == "1":
                        first_cell = sheet.cell(i, j)
                        print(sheet.title, i, j)  # debug

    wb = openpyxl.load_workbook(str(file_path))
    print(wb.sheetnames)
    for sheet_name in wb.sheetnames:
        for name, system in sheet_names.items():
            if name in sheet_name:
                parse_sheet(wb[sheet_name], system)


if __name__ == "__main__":
    jobs_schedule_asu = Path(r"c:\Users\Mihail\Documents\Гусев М.В\Планы работ\Графики ТО\5. Графики на 05.18 АСУ.XLSX")
    parser_asu(jobs_schedule_asu)
