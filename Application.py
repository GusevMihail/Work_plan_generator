from os import listdir
from itertools import groupby
from typing import List
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


def parse_asu_folder():
    jobs_list: List[Pre_processing.Job] = []
    path_asu = r'.\input data\АСУ'
    for file in get_xlsx_files(path_asu):
        print(f'parse asu file: {file}')
        file_path = path_asu + '\\' + str(file)
        workbook_asu = openpyxl.load_workbook(file_path)
        for sheet_name in workbook_asu.sheetnames:  # find necessary worksheets by names
            system = Pre_processing.extract_system(sheet_name)
            if system is not None:
                parser = Parser.ParserAsu(workbook_asu[sheet_name])
                Pre_processing.parser_to_jobs(parser, jobs_list)


if __name__ == "__main__":
    # workbook_asu = openpyxl.load_workbook(r'.\input data\2. АСУ 02.19.xlsx')
    # workbook_asu = openpyxl.load_workbook(r'.\input data\Test Schedule.xlsx')
    jobs = []
    jobs.extend(parse_asu_folder())

    # jobs.sort(key=lambda job: (job.date, job.object, job.system, job.work_type))
    # jobs_by_days = groupby(jobs, key=lambda job: job.date)
    # for j in jobs_by_days:
    #     template_filename = r'.\input data\Template.xlsx'
    #     day_job = list(j[1])
    #     test_table = Table_generator.WorkPlan(day_job, template_filename)
    #     test_table.make_plan()
    #     test_table.save_file()
    #
    # print('Генерация успешно завершена')

    # workbook_vols = openpyxl.load_workbook(r'.\input data\05 май ВОЛС.xlsx')
    # sheet = workbook_vols['10.4.38 ТО']
    # parserVOLS = Parser.ParserVOLS(sheet)
    # parser._find_place()

    # parserVOLS._find_data_boundaries()
    # parserVOLS._find_month_year()
    # parserVOLS._extract_jobs()
