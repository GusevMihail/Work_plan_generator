from collections import namedtuple
from typing import List
from abc import ABCMeta, abstractmethod, abstractproperty

from openpyxl.worksheet import Worksheet

from Cell_styler import TableArea


def xstr(cell_value):
    if cell_value is None:
        return None
    else:
        return str(cell_value)


def xint(cell_value):
    if cell_value is None:
        return None
    else:
        return int(cell_value)


RawData = namedtuple('raw_data', 'day work_type place')


class AbstractParser(metaclass=ABCMeta):

    def __init__(self, sheet: Worksheet):
        self.sheet = sheet
        self.month_year = None
        self.raw_data: List[RawData] = []

    @abstractmethod
    def _find_data_boundaries(self):
        raise NotImplementedError

    @abstractmethod
    def _find_month_year(self):
        raise NotImplementedError

    @abstractmethod
    def _extract_jobs(self):
        raise NotImplementedError


class ParserAsu(AbstractParser):

    def __init__(self, sheet: Worksheet):
        super().__init__(sheet)
        self._find_data_boundaries()
        self._find_month_year()
        self._extract_jobs()
        self._data_area = None

    def _find_data_boundaries(self):
        max_table_row = 500
        max_table_col = 60
        first_row = None
        first_col = None
        last_row = None
        last_col = None

        for row in range(1, max_table_row):
            if xstr(self.sheet.cell(row, 1).value) == '1':
                for col in range(1, max_table_col):
                    if xstr(self.sheet.cell(row - 1, col).value) == '1':
                        first_row = row
                        first_col = col
                        break
                break

        for row in range(first_row, max_table_row):
            object_name_col = 2
            cell = self.sheet.cell(row, object_name_col).value
            if not isinstance(cell, str):
                last_row = row - 1
                break

        for col in range(first_col, max_table_col):
            if self.sheet.cell(first_row - 1, col).value is None:
                last_col = col - 1
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
            place_col = 2
            raw_place = self.sheet.cell(i_row, place_col).value
            for i_col in range(self._data_area.first_col, self._data_area.last_col + 1):
                raw_work_type = self.sheet.cell(i_row, i_col).value
                if raw_work_type is not None:
                    raw_day = self.sheet.cell(self._data_area.first_row - 1, i_col).value
                    i_raw_data = RawData(raw_day, raw_work_type, raw_place)
                    self.raw_data.append(i_raw_data)
