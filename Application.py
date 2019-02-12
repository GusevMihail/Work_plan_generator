from os import listdir
from itertools import groupby
from typing import List, Tuple
import openpyxl

import Parser
import Pre_processing
import Table_generator


# TODO перенести обработку отдельных типов файлов в отдельные функции. Брать файлы из прописанных папок (АСУ, ВОЛС).

def get_xlsx_files(path):
    files = listdir(path)
    xlsx_files = filter(lambda x: '.xlsx' in x and '$' not in x, files)
    print('run')
    return xlsx_files


def find_sheets_asu(wb: openpyxl.Workbook) -> Tuple[openpyxl.workbook.workbook.Worksheet]:
    return tuple(sheet for sheet in wb.worksheets
                 if Pre_processing.extract_system(sheet.title) is not None and sheet.sheet_state == 'visible')


def find_sheets_vols(wb: openpyxl.Workbook) -> Tuple[openpyxl.workbook.workbook.Worksheet]:
    return tuple(sheet for sheet in wb.worksheets if 'ТО' in sheet.title and sheet.sheet_state == 'visible')


def process_files(folder: str, find_sheets_function, parser_class):
    jobs_list: List[Pre_processing.Job] = []
    for file in get_xlsx_files(folder):
        print(f'parse asu file: {file}')
        file_path = folder + '\\' + str(file)
        workbook = openpyxl.load_workbook(file_path)
        sheets = find_sheets_function(openpyxl.load_workbook(file_path))
        for sheet in sheets:
            parser = parser_class(sheet)
            jobs_list.extend(Pre_processing.parser_to_jobs(parser))
    return jobs_list


def make_xlsx_from_jobs(jobs_list):
    jobs.sort(key=lambda job: (job.date, job.object, job.system, job.work_type))
    jobs_by_days = groupby(jobs, key=lambda job: job.date)
    for j in jobs_by_days:
        template_filename = r'.\input data\Template.xlsx'
        day_job = list(j[1])
        test_table = Table_generator.WorkPlan(day_job, template_filename)
        test_table.make_plan()
        test_table.save_file()


if __name__ == "__main__":
    workbook_asu = openpyxl.load_workbook(r'.\input data\2. АСУ 02.19.xlsx')
    workbook_test_vols = openpyxl.load_workbook(r'.\input data\Test Schedule VOLS.xlsx')

    jobs = []
    jobs.extend(process_files(r'.\input data\АСУ', find_sheets_asu, Parser.ParserAsu))
    print(len(jobs))

    # print(find_sheets_vols(workbook_test_vols))

    # make_xlsx_from_jobs(jobs)
    print('Генерация успешно завершена')

    # workbook_vols = openpyxl.load_workbook(r'.\input data\05 май ВОЛС.xlsx')
    # sheet = workbook_vols['10.4.38 ТО']
    # parserVOLS = Parser.ParserVOLS(sheet)
    # parser._find_place()

    # parserVOLS._find_data_boundaries()
    # parserVOLS._find_month_year()
    # parserVOLS._extract_jobs()
