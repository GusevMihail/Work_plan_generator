from itertools import groupby

import openpyxl

import Parser
import Pre_processing
import Table_generator

if __name__ == "__main__":
    workbook = openpyxl.load_workbook(r'.\input data\5. Графики на 05.18 АСУ.xlsx')
    # workbook = openpyxl.load_workbook(r'.\input data\Test Schedule.xlsx')
    jobs = []
    for sheet_name in workbook.sheetnames:  # find necessary worksheets by names
        system = Pre_processing.extract_system(sheet_name)
        if system is not None:
            parser = Parser.ParserAsu(workbook[sheet_name])
            Pre_processing.parser_to_jobs(parser, jobs)
    jobs.sort(key=lambda job: (job.date, job.object, job.system, job.work_type))

    # for j in jobs:  # debug
    #     print(j)

    jobs_by_days = groupby(jobs, key=lambda job: job.date)
    for j in jobs_by_days:
        template_filename = r'.\input data\Template.xlsx'
        day_job = list(j[1])
        test_table = Table_generator.WorkPlan(day_job, template_filename)
        test_table.make_plan()
        test_table.save_file()

    print('Генерация успешно завершена')
