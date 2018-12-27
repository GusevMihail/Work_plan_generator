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
    month = None
    for m_name, m_num in month_names.items():
        if m_name in raw_date.lower():
            month = m_num
    year = int(re.findall(r'\d+', raw_date)[0])
    # print(raw_date, month, year)  # debug
    return month, year


def extract_place(raw_place: str):
    places_names = {'ЗУ КЗС': ('Здание управления КЗС', 'ЗУ'),
                    'Здание управления': ('Здание управления КЗС', 'ЗУ'),
                    'АМ': ('С2 АМ', 'С2'),
                    }

    raw_place = raw_place.strip(' ,.\t\n')
    raw_place = str.replace(raw_place, 'c', 'с')    # Eng to Rus
    raw_place = str.replace(raw_place, 'C', 'С')    # Eng to Rus
    raw_place = str.replace(raw_place, 'ВЗ', 'В3')  # Letter to Num
    raw_place = str.replace(raw_place, 'север', 'Север')
    raw_place = str.replace(raw_place, 'юг', 'Юг')
    raw_place = str.replace(raw_place, '(', '')
    raw_place = str.replace(raw_place, ')', '')

    for i_template, i_place in places_names.items():
        if i_template in raw_place:
            return i_place

    # find В1..В6 objects
    search_obj = re.search(r'В\W{,3}(\d)', raw_place)
    if search_obj:
        return 'В' + search_obj.group(1), 'В' + search_obj.group(1)

    # find С1, С2 objects
    search_obj = re.search(r'(С\d)(.*)', raw_place)
    if search_obj:
        return ''.join(search_obj.groups()), search_obj.group(1)
    else:
        print('нет совпадений с шаблоном')  # debug
        return raw_place, 'unknown'


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

        end_data_row = origin.row
        for i_row in range(origin.row, origin.row + 20):
            if sheet.cell(i_row, 1).value is None:
                end_data_row: int = i_row - 1
                break
        print(end_data_row)

        end_data_col = origin.col_idx
        for i_col in range(origin.col_idx, origin.col_idx + 40):
            if sheet.cell(origin.row - 1, i_col).value is None:
                end_data_col = i_col - 1
                break
        print(f' end data col {end_data_col}; last date {sheet.cell(origin.row - 1, end_data_col).value}')

        # for i_row in range(origin.row, end_data_row + 1):
        #     place = sheet.cell(i_row, 2).value
        #     print(place)

    wb = openpyxl.load_workbook(str(file_path))
    for sheet_name in wb.sheetnames:  # find necessary worksheets by names
        for name, system in sheet_names.items():
            if name in sheet_name:
                parse_sheet(wb[sheet_name], system)
                break


if __name__ == "__main__":
    jobs_schedule_asu = Path(
        r"c:\Users\Mihail\PycharmProjects\Work_plan_generator\input data\\5. Графики на 05.18 АСУ.xlsx")
    # parser_asu(jobs_schedule_asu)

    test_raw_places = open('test raw places.txt')
    for line in test_raw_places:
        print(f'{line}  -->>  {extract_place(line)}')
    # print(extract_place('Судопропускное сооружение Са1 Юг ДКФ'))
