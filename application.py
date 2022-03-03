from itertools import groupby
from os import listdir
from typing import List, Tuple, Union, Any

import openpyxl

import config_email
import duty_schedule
import email_processing
import table_generator
import works_parser
from config_journals import batch_ASU_journals, batch_ASKUE_journals
from journal import JournalASU, JournalASKUE, jobs_to_df, batch_journal_generator
from pre_processing import find_system_by_sheet, Job, parser_to_jobs
from work_calendar import batch_make_calendars


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
    print('Генерация успешно завершена \n')


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
    # input()

    # for j in jobs:
    #     print(j)

    df_jobs = jobs_to_df(jobs)

    print('Генерация журналов работ')
    batch_journal_generator(df_jobs, JournalASU, batch_ASU_journals, verbose=True)
    batch_journal_generator(df_jobs, JournalASKUE, batch_ASKUE_journals, verbose=True)
    print('Генерация успешно завершена \n')

    print('Генерация календарей работ')
    batch_make_calendars(df_jobs, verbose=True)
    print('Генерация успешно завершена \n')

    is_sand_plans, is_sand_journals, is_sand_calendars = [False] * 3
    while (answer := input('send all files? [y]/n\n')) not in ('y', 'n', ''):
        pass
    else:
        if answer == 'y' or '':
            is_sand_plans, is_sand_journals, is_sand_calendars = [True] * 3
        else:
            while (answer := input('send plans? [y]/n\n')) not in ('y', 'n', ''):
                pass
            else:
                is_sand_plans = (answer == 'y' or '')

            while (answer := input('send journals? [y]/n\n')) not in ('y', 'n', ''):
                pass
            else:
                is_sand_journals = (answer == 'y' or '')

            while (answer := input('send calendars? [y]/n\n')) not in ('y', 'n', ''):
                pass
            else:
                is_sand_calendars = (answer == 'y' or '')

    if is_sand_plans:
        email_processing.send_journals(config_email.batch_sending_plans,
                                       attachment_folder=r'./output data/plans/',
                                       mail_subj='планы работ',
                                       add_month_to_subj=True, test_mod=False)
    if is_sand_journals:
        email_processing.send_journals(config_email.batch_sending_journals,
                                       attachment_folder=r'./output data/journals/',
                                       mail_subj='журналы работ',
                                       add_month_to_subj=True, test_mod=False)
    if is_sand_calendars:
        email_processing.send_journals(config_email.batch_sending_calendars,
                                       attachment_folder=r'./output data/calendars/',
                                       mail_subj='календарь работ',
                                       add_month_to_subj=True, test_mod=False)
