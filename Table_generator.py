from datetime import date
import os  # debug
from typing import List

import openpyxl
from openpyxl.styles import Font, Border, Side, Alignment
from openpyxl.worksheet import Worksheet

from Cell_styler import apply_style, TableArea
from Pre_processing import Job


class WorkPlan:

    def __init__(self, jobs: List[Job], template_path: str):
        self.jobs = jobs
        self._first_col = None
        self._last_col = None
        self._first_data_row = None
        self._current_row = None
        self._thin_side = Side(style='thin', color='000000')
        self._thin_border = Border(left=self._thin_side, top=self._thin_side,
                                   right=self._thin_side, bottom=self._thin_side)
        self._basic_font = Font(name='Times New Roman', size=11)
        self._bold_font = Font(name='Times New Roman', size=11, b=True)
        self._align_center = Alignment(horizontal='center', vertical='center')
        self._wb = self._get_template(template_path)
        self._ws = self._wb.active
        self._write_date()

    def _get_template(self, filename: str):
        wb_template = openpyxl.load_workbook(filename)
        ws_template = wb_template.active
        self._first_col = 1
        self._last_col = 9
        self._first_data_row = 10
        self._current_row = self._first_data_row
        table_header = TableArea(first_row=8, last_row=9, first_col=self._first_col, last_col=self._last_col)
        apply_style(ws_template, table_header, border=self._thin_border)
        return wb_template

    def _write_date(self):
        date_cell = self._ws.cell(row=6, column=1)
        date_cell.value = self.jobs[0].date

    def _write_obj_row(self, obj_name):
        self._ws.merge_cells(start_row=self._current_row, end_row=self._current_row,
                             start_column=self._first_col, end_column=self._last_col)

        current_row_area = TableArea(self._current_row, self._first_col, self._current_row, self._last_col)
        apply_style(self._ws, current_row_area, border=self._thin_border, font=self._bold_font)
        self._ws.cell(self._current_row, self._first_col).value = obj_name
        self._current_row += 1

    def _write_work_row(self, job: Job):
        current_row_area = TableArea(self._current_row, self._first_col, self._current_row, self._last_col)
        apply_style(self._ws, current_row_area,
                    border=self._thin_border,
                    font=self._basic_font,
                    alignment=self._align_center)
        organization_col = 1
        organization = 'Би. Си. Си. ООО'
        system_col = 2
        work_col = 3
        place_col = 4
        work_start_col = 5
        work_start = '9:00'
        work_end_col = 6
        work_end = '18:00'
        worker_col = 7
        job.worker = 'Подольский Андрей Вениаминович +79522726122'  # temporary code for test
        self._ws.cell(self._current_row, organization_col).value = organization
        self._ws.cell(self._current_row, system_col).value = job.system
        self._ws.cell(self._current_row, work_col).value = job.work_type
        self._ws.cell(self._current_row, place_col).value = job.place
        self._ws.cell(self._current_row, work_start_col).value = work_start
        self._ws.cell(self._current_row, work_end_col).value = work_end
        self._ws.cell(self._current_row, worker_col).value = job.worker

    def make_plan(self):
        for job in self.jobs:
            self._write_work_row(job)

    def save_file(self, filename):
        self._wb.save(filename)


if __name__ == '__main__':
    # test_out_wb = openpyxl.Workbook()
    # template_filename = r'.\input data\Template.xlsx'
    # test_out_wb = get_template(template_filename)
    # test_ws = test_out_wb.worksheets[0]
    # test_date = date(2019, 1, 20)
    # write_date(test_ws, test_date)
    # write_obj_row(test_ws, 'TEST')

    # style1 = NamedStyle(name='Style1')
    # style1.border = Border(left=thin_side, top=thin_side, right=thin_side, bottom=thin_side)
    # style1.alignment = Alignment(vertical='center', horizontal='center')
    # style1.fill = PatternFill(fill_type='solid', start_color='ff8327', end_color='ff8327')
    # test_out_wb.add_named_style(style1)

    dest_filename = r'test_wb.xlsx'
    # test_out_wb.save(filename=dest_filename)
    # os.startfile(dest_filename)  # debug
