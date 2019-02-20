from itertools import groupby
from os import listdir
from typing import List, Tuple

import openpyxl

import Parser
import Pre_processing
import Table_generator


def get_xlsx_files(path):
    files = listdir(path)
    xlsx_files = filter(lambda x: '.xlsx' in x and '$' not in x, files)
    print('run')
    return xlsx_files


def find_sheets_asu(wb: openpyxl.Workbook) -> Tuple[openpyxl.workbook.workbook.Worksheet]:
    return tuple(sheet for sheet in wb.worksheets
                 if Pre_processing.find_system_by_sheet(sheet.title) is not None and sheet.sheet_state == 'visible')


def find_sheets_vols(wb: openpyxl.Workbook) -> Tuple[openpyxl.workbook.workbook.Worksheet]:
    return tuple(sheet for sheet in wb.worksheets if 'ТО' in sheet.title and sheet.sheet_state == 'visible')


def process_files(folder: str, find_sheets_function, parser_class):
    jobs_list: List[Pre_processing.Job] = []
    for file in get_xlsx_files(folder):
        print(f'parse asu file: {file}')
        file_path = folder + '\\' + str(file)
        workbook = openpyxl.load_workbook(file_path)
        sheets = find_sheets_function(workbook)
        for sheet in sheets:
            parser = parser_class(sheet)
            jobs_list.extend(Pre_processing.parser_to_jobs(parser))
    return jobs_list


def make_xlsx_from_jobs(jobs_list):
    jobs_list.sort(key=lambda x: (x.date, x.object, x.system, x.work_type))
    jobs_by_days = groupby(jobs, key=lambda x: x.date)
    for job in jobs_by_days:
        template_filename = r'.\input data\Template.xlsx'
        day_job = list(job[1])
        table = Table_generator.WorkPlan(day_job, template_filename)
        table.make_plan()
        table.save_file()


if __name__ == "__main__":
    jobs = []
    jobs.extend(process_files(r'.\input data\АСУ', find_sheets_asu, Parser.ParserAsu))
    jobs.extend(process_files(r'.\input data\ВОЛС', find_sheets_vols, Parser.ParserVols))
    jobs.extend(process_files(r'.\input data\Телеканал', find_sheets_vols, Parser.ParserTk))
    jobs.extend(process_files(r'.\input data\АИИСКУЭ', find_sheets_vols, Parser.ParserAskue))
    jobs.extend(process_files(r'.\input data\Тех.учет', find_sheets_vols, Parser.ParserTechReg))
    make_xlsx_from_jobs(jobs)
    print('Генерация успешно завершена')
    print(f'Всего найдено работ: {len(jobs)}')

    # for j in jobs:
    #     print(j)
