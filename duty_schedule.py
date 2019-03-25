from openpyxl.utils import get_column_letter
from openpyxl.worksheet import Worksheet
from parser import xint


class DutySchedule:
    def __init__(self, worksheet: Worksheet):
        self.worksheet = worksheet
        self._data_first_col = 1
        self._data_last_col = None
        self._find_last_day_col()
        # номера строк вносятся в порядке убывания приоритета сотрудника для выбора его в качестве руководителя работ
        self.group_s1_rows = (8, 9, 10, 11)
        self.group_s2_rows = (14, 13)
        self.group_v_rows = (17, 18, 19)
        self.group_vols_rows = (23, 21, 22)
        self.group_tk_rows = (22, 21, 23)

    def _find_last_day_col(self):
        days_row = 2
        table_max_col = 50
        for col in range(self._data_first_col, table_max_col + 1):
            cell = self.worksheet.cell(days_row, col)
            if xint(cell.value) is not int:
                self._data_last_col = col - 1
        else:
            print('Последняя дата не найдена')

    # def is_workday(self, column:int) -> bool:
    #     for cell in
