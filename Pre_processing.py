import datetime
import re
from typing import List
import random

import Parser


class Job:
    def __init__(self):
        self.date = None
        self.object = None
        self.system = None
        self.work_type = None
        self.place = None
        self.worker = None

    @staticmethod
    def _print_str(value, length=0):
        if value is None:
            return '<None>'.ljust(length)
        else:
            return str(value).ljust(length)

    def __str__(self):
        return f'obj:{self._print_str(self.object, 35)}' \
            f'place:{self._print_str(self.place, 40)}' \
            f'work:{self._print_str(self.work_type, 10)}' \
            f'date:{self._print_str(self.date, 15)}' \
            f'sys:{self._print_str(self.system, 15)}' \
            f'worker:{self._print_str(self.worker)}'

    def __repr__(self):
        return f'obj:{self.object}; place:{self.place}; work:{self.work_type}; ' \
            f'date:{self.date};  sys:{self.system}; worker:{self.worker}'

    def find_worker(self):
        group_s1 = ('Гусев Михаил Владимирович +79675904368',
                    'Харитонов Виктор Яковлевич +79319642368',
                    'Мулин Николай Николаевич +79112283889',
                    'Кушмылев Евгений Павлович +79819134737')

        group_s2 = ('Каприца Анатолий Евгеньевич +79046315018',
                    # 'Макаров Виктор Викторович +79312089129',
                    'Горнов Александр Серафимович +79111314015')

        group_v = ('Подольский Андрей Вениаминович +79312531066',
                   'Кокоев Михаил Николаевич +79216441993',
                   'Санжара Владимир Александрович +79819629285')

        group_vols = ('Ильин Андрей Владимирович +79219303652',
                      'Ястребов Алексей Владимирович +79313581975',
                      'Огородников Алексей Юрьевич +79313196196')

        random.seed(self.date)
        if self.system in ('ЛВС', 'ВОЛС', 'Телеканал М2'):
            self.worker = random.choice(group_vols)
        elif 'С1' in self.object:
            self.worker = random.choice(group_s1)
        elif 'С2' in self.object:
            self.worker = random.choice(group_s2)
        else:
            self.worker = random.choice(group_v)


def find_num_in_str(string: str) -> int:
    result = re.findall(r'\d+', string)
    return int(result[0])


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
    places_names = {'ЗУ КЗС': ('Здание управления КЗС', 'Здание управления КЗС'),
                    'Здание управления': ('Здание управления КЗС', 'Здание управления КЗС'),
                    'АМ': ('С2 АМ', 'Судопропускное сооружение С2'),
                    }

    raw_place = raw_place.strip(' ,.\t\n')
    raw_place = raw_place.replace('c', 'с')  # Eng to Rus
    raw_place = raw_place.replace('C', 'С')  # Eng to Rus
    raw_place = raw_place.replace('ВЗ', 'В3')  # Letter to Num
    raw_place = raw_place.replace('С-', 'С')  # C-1\C-2 to C1\C2
    raw_place = raw_place.replace('В-', 'В')  # В-№ to В№
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
        return 'В' + search_obj.group(1), 'Водопропускное сооружение ' + 'В' + search_obj.group(1)

    # find PS 86
    if 'ПС' in raw_place and '86' in raw_place:
        return 'ПС 86', 'Судопропускное сооружение С1'

    # find PS 360
    if 'ПС' in raw_place and '360' in raw_place:
        return 'ПС 360', 'Горская'

    # find PS 223
    if 'ПС' in raw_place and '223' in raw_place:
        return 'ПС 223', 'Бронка'

    # find PS S1 110/10
    if 'ПС' in raw_place and 'С1' in raw_place and '110' in raw_place:
        return 'С1 ПС 110/10кВ', 'Судопропускное сооружение С1'

    # find PS S2 110/10
    if 'ПС' in raw_place and 'С2' in raw_place and '110' in raw_place:
        return 'С2 ПС 110/10кВ', 'Судопропускное сооружение С2'

    # find С1, С2 objects
    search_obj = re.search(r'(С\d)(.*)', raw_place)
    if search_obj:
        return ''.join(search_obj.groups()), 'Судопропускное сооружение ' + search_obj.group(1)

    print(f'extract_place_and_object: {raw_place} - нет совпадений с шаблоном')  # debug
    return raw_place, 'unknown'


def find_system_by_sheet(sheet_name):
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


def parser_to_jobs(parser) -> List[Job]:
    jobs: List[Job] = []
    month, year = extract_month_and_year(parser.month_year)
    # system = find_system_by_sheet(parser.sheet.title)
    for raw_job in parser.raw_data:
        job = Job()
        job.place, job.object = extract_place_and_object(raw_job.place)
        job.work_type = filter_work_type(raw_job.work_type)
        job.date = datetime.date(year, month, raw_job.day)
        job.system = parser.system
        job.find_worker()
        jobs.append(job)
    return jobs
