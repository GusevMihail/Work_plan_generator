import datetime
import re
from typing import List

import Parser


class Job:
    def __init__(self):
        self.date = None
        self.object = None
        self.system = None
        self.work_type = None
        self.place = None
        self.worker = None

    def __str__(self):
        return f'obj:{self.object.ljust(5)} place:{self.place.ljust(30)} work:{self.work_type.ljust(10)} date:{self.date}  sys:{self.system}'


def extract_month_and_year(raw_date: str):
    month_names = {'январь': 1,
                   'февраль': 2,
                   'март': 3,
                   'апрель': 4,
                   'май': 5,
                   'июнь': 6,
                   'июль': 7,
                   'август': 8,
                   'сентябрь': 9,
                   'октябрь': 10,
                   'ноябрь': 11,
                   'декабрь': 12}
    month = None
    for m_name, m_num in month_names.items():
        if m_name in raw_date.lower():
            month = m_num
    year = int(re.findall(r'\d+', raw_date)[0])
    # print(raw_date, month, year)  # debug
    return month, year


def extract_place_and_object(raw_place: str):
    places_names = {'ЗУ КЗС': ('Здание управления КЗС', 'ЗУ'),
                    'Здание управления': ('Здание управления КЗС', 'ЗУ'),
                    'АМ': ('С2 АМ', 'С2'),
                    }

    raw_place = raw_place.strip(' ,.\t\n')
    raw_place = raw_place.replace('c', 'с')  # Eng to Rus
    raw_place = raw_place.replace('C', 'С')  # Eng to Rus
    raw_place = raw_place.replace('ВЗ', 'В3')  # Letter to Num
    raw_place = raw_place.replace('север', 'Север')
    raw_place = raw_place.replace('юг', 'Юг')
    raw_place = raw_place.replace('(', '')
    raw_place = raw_place.replace(')', '')

    for i_template, i_place in places_names.items():
        if i_template in raw_place:
            return i_place

    # find В1..В6 objects
    search_obj = re.search(r'В\W{,3}(\d)', raw_place)
    if search_obj:
        return 'В' + search_obj.group(1), 'В' + search_obj.group(1)

    # find С1, С2 objects
    search_obj = re.search(r'(С\d)(.*)', raw_place)
    if search_obj:
        return ''.join(search_obj.groups()), search_obj.group(1)
    else:
        print('нет совпадений с шаблоном')  # debug
        return raw_place, 'unknown'


def extract_system(sheet_name):
    sheet_names = {'АСУ ТП': 'АСУ ТП',
                   'АСУ И': 'АСУ И',
                   'МОСТ': 'АСУ АМ',
                   'ЛВС': 'ЛВС'}

    for name, system in sheet_names.items():
        if name in sheet_name:
            return system
    return None


def filter_work_type(work: str):
    # work = work.strip(' ,.\t\n')
    work = work.replace('E', 'Е')  # Eng to Rus
    work = work.replace('T', 'Т')  # Eng to Rus
    work = work.replace('O', 'О')  # Eng to Rus
    work = work.replace('ТОЗ', 'ТО3')  # Letter to Num
    return work


def parser_to_jobs(parser: Parser.ParserAsu, jobs: List[Job]):
    month, year = extract_month_and_year(parser.month_year)
    system = extract_system(parser.sheet.title)
    for raw_job in parser.raw_data:
        job = Job()
        job.place, job.object = extract_place_and_object(raw_job.place)
        job.work_type = filter_work_type(raw_job.work_type)
        job.date = datetime.date(year, month, raw_job.day)
        job.system = system
        jobs.append(job)
        # print(raw_job) # debug
        # print(f'obj {job.object}\t pl {job.place.ljust(30)} wt {job.work_type.ljust(10)} dt {job.date}\t sys {job.system}') # debug
