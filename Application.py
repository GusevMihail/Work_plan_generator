import os

import openpyxl

import Parser
import Pre_processing
import Table_generator

if __name__ == "__main__":
    # workbook = openpyxl.load_workbook(r'.\input data\5. Графики на 05.18 АСУ.xlsx')
    workbook = openpyxl.load_workbook(r'.\input data\Test Schedule.xlsx')
    jobs = []
    for sheet_name in workbook.sheetnames:  # find necessary worksheets by names
        system = Pre_processing.extract_system(sheet_name)
        if system is not None:
            parser = Parser.ParserAsu(workbook[sheet_name])
            Pre_processing.parser_to_jobs(parser, jobs)
    # for job in jobs:
    #     print(job)
    template_filename = r'.\input data\Template.xlsx'
    output_filename = r'test_wb.xlsx'
    test_table = Table_generator.WorkPlan(jobs, template_filename)
    test_table._write_obj_row('somwhere in far far away damba')
    test_table.make_plan()
    test_table.save_file(output_filename)

    os.startfile(output_filename)  # debug
