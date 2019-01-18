# from pathlib import Path
from collections import namedtuple

# import openpyxl
from openpyxl.worksheet import Worksheet


# from Application import Job
# from Pre_processing import extract_month_and_year


# class RawData:
#     def __init__(self):
#         self.day = None
#         # self.object = None
#         # self.system = None
#         self.work_type = None
#         self.place = None
#         # self.sheet_name = None


def xstr(cell_value):
    if cell_value is None:
        return None
    else:
        return str(cell_value)


raw_data = namedtuple('raw_data', 'day work_type place')
table_area = namedtuple('table_area', 'first_row first_col last_row last_col')


class ParserAsu:

    def __init__(self, sheet: Worksheet):
        self.sheet = sheet
        self.data_area = None
        self.month_year = None
        self.raw_data = []

    def find_data_boundaries(self):
        first_row = 1
        first_col = 1
        for i_row in range(1, 30):
            if xstr(self.sheet.cell(i_row, 1).value) == "1":
                for i_col in range(1, 30):
                    if xstr(self.sheet.cell(i_row - 1, i_col).value) == "1":
                        first_row = i_row
                        first_col = i_col
                        break
                break
        last_row = first_row
        for i_row in range(first_row, first_row + 20):
            if self.sheet.cell(i_row, 1).value is None:
                last_row = i_row - 1
                break
        last_col = first_col
        for i_col in range(first_col, first_col + 40):
            if self.sheet.cell(first_row - 1, i_col).value is None:
                last_col = i_col - 1
                break
        self.data_area = table_area(first_row, first_col, last_row, last_col)

    def find_month_year(self):
        self.month_year = self.sheet.cell(self.data_area.first_row - 2, self.data_area.first_col).value
        if self.month_year is None:
            self.month_year = self.sheet.cell(self.data_area.first_row - 3, self.data_area.first_col).value

    def extract_jobs(self):
        for i_row in range(self.data_area.first_row, self.data_area.last_row + 1):
            raw_place = self.sheet.cell(i_row, 2).value
            for i_col in range(self.data_area.first_col, self.data_area.last_col + 1):
                # i_raw_data.place = raw_place
                raw_work_type = self.sheet.cell(i_row, i_col).value
                if raw_work_type is not None:
                    raw_day = self.sheet.cell(self.data_area.first_row - 1, i_col).value
                    i_raw_data = raw_data(raw_day, raw_work_type, raw_place)
                    print(i_raw_data)  # debug
                    self.raw_data.append(i_raw_data)
