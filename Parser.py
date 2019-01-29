from collections import namedtuple
from typing import List

from openpyxl.worksheet import Worksheet

from Cell_styler import TableArea


def xstr(cell_value):
    if cell_value is None:
        return None
    else:
        return str(cell_value)


RawData = namedtuple('raw_data', 'day work_type place')


class ParserAsu:

    def __init__(self, sheet: Worksheet):
        self.sheet = sheet
        self.month_year = None
        self.raw_data: List[RawData] = []
        self._data_area = None

        self._find_data_boundaries()
        self._find_month_year()
        self._extract_jobs()

    def _find_data_boundaries(self):
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
        self._data_area = TableArea(first_row, first_col, last_row, last_col)

    def _find_month_year(self):
        if self._data_area is None:
            self._find_data_boundaries()
        self.month_year = self.sheet.cell(self._data_area.first_row - 2, self._data_area.first_col).value
        if self.month_year is None:
            self.month_year = self.sheet.cell(self._data_area.first_row - 3, self._data_area.first_col).value

    def _extract_jobs(self):
        if self._data_area is None:
            self._find_data_boundaries()
        for i_row in range(self._data_area.first_row, self._data_area.last_row + 1):
            raw_place = self.sheet.cell(i_row, 2).value
            for i_col in range(self._data_area.first_col, self._data_area.last_col + 1):
                raw_work_type = self.sheet.cell(i_row, i_col).value
                if raw_work_type is not None:
                    raw_day = self.sheet.cell(self._data_area.first_row - 1, i_col).value
                    i_raw_data = RawData(raw_day, raw_work_type, raw_place)
                    # print(i_raw_data)  # debug
                    self.raw_data.append(i_raw_data)
