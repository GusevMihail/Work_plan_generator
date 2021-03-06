from itertools import groupby
from os import listdir
from typing import List, Tuple, Union, Any

import openpyxl

import table_generator
import works_parser
import duty_schedule
from pre_processing import find_system_by_sheet, Job, parser_to_jobs


def get_xlsx_files(path):
    files = listdir(path)
    xlsx_files = filter(lambda x: '.xlsx' in x and '$' not in x, files)
    return xlsx_files


def find_sheets_asu(wb: openpyxl.Workbook) -> Union[Tuple[openpyxl.workbook.workbook.Worksheet], Any]:
    return tuple(sheet for sheet in wb.worksheets
                 if find_system_by_sheet(sheet.title) is not None and sheet.sheet_state == 'visible')


def find_sheets_vols(wb: openpyxl.Workbook) -> Union[Tuple[openpyxl.workbook.workbook.Worksheet], Any]:
    return tuple(sheet for sheet in wb.worksheets if 'ТО' in sheet.title and sheet.sheet_state == 'visible')


def all_visible_sheets(wb: openpyxl.Workbook) -> Union[Tuple[openpyxl.workbook.workbook.Worksheet], Any]:
    return tuple(sheet for sheet in wb.worksheets if sheet.sheet_state == 'visible')


def process_files(folder: str, find_sheets_function, parser_class) -> List[Job]:
    jobs_list = []
    print(f'folder: {folder}')
    for file in get_xlsx_files(folder):
        print(f' - {file}')
        file_path = folder + '\\' + str(file)
        workbook = openpyxl.load_workbook(file_path)
        sheets = find_sheets_function(workbook)
        for sheet in sheets:
            sheet_parser = parser_class(sheet)
            jobs_list.extend(parser_to_jobs(sheet_parser))
    return jobs_list


def process_duty_schedules(folder: str) -> Tuple[duty_schedule.DutySchedule]:
    schedules = []
    for file in get_xlsx_files(folder):
        file_path = folder + '\\' + str(file)
        ws = openpyxl.load_workbook(file_path).worksheets[0]
        schedules.append(duty_schedule.DutySchedule(ws, duty_schedule.all_workers))
        return tuple(schedules)


def make_xlsx_from_jobs(jobs_list: List[Job]):
    print('Генерация планов работ')
    jobs_list.sort(key=lambda x: (x.date, x.object.value, x.system.value, x.work_type))

    jobs_by_days = groupby(jobs_list, key=lambda x: x.date)
    for day_jobs in jobs_by_days:
        template_filename = r'.\input data\Template.xlsx'
        day_jobs = list(day_jobs[1])  # распаковка результата группировки
        # отбрасывание элементов списка, не уникальных по комбинации некоторых полей.
        day_jobs = list({(j.object, j.place, j.system, j.work_type): j for j in day_jobs}.values())

        table = table_generator.WorkPlan(day_jobs, template_filename)
        table.make_plan()
        table.save_file()
    print('Генерация успешно завершена')


duty_schedules = process_duty_schedules(r'.\input data\Графики дежурств')

if __name__ == "__main__":
    jobs = []
    jobs.extend(process_files(r'.\input data\SAKE', all_visible_sheets, works_parser.ParserSake))  # all systems
    # jobs.extend(process_files(r'.\input data\1', all_visible_sheets, works_parser.ParserSake))  # tests
    # jobs.extend(process_files(r'.\input data\АСУ', find_sheets_asu, works_parser.ParserAsu))
    # jobs.extend(process_files(r'.\input data\ВОЛС', all_visible_sheets, works_parser.ParserVols_v2))
    # jobs.extend(process_files(r'.\input data\Телеканал', find_sheets_vols, works_parser.ParserTk))
    # jobs.extend(process_files(r'.\input data\АИИСКУЭ', find_sheets_vols, works_parser.ParserAskueSake))
    # jobs.extend(process_files(r'.\input data\Тех.учет', find_sheets_vols, works_parser.ParserTechReg))
    print(f'Всего найдено работ: {len(jobs)}')
    make_xlsx_from_jobs(jobs)

    # print(f'Всего найдено работ: {len(jobs)}')
    input()

    for j in jobs:
        print(j)
