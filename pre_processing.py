import datetime
import random
import re
from enum import Enum
from typing import List, Optional


class Systems(Enum):
    ASU_TP = 'АСУ ТП'
    ASU_I = 'АСУ И'
    ASU_AM = 'АСУ АМ'
    LVS = 'ЛВС'
    VOLS = 'ВОЛС'
    TK = 'Телеканал М2'
    ASKUE = 'АИИСКУЭ'
    TECH_REG = 'Тех. учет'


class Objects(Enum):
    S1 = 'Судопропускное сооружение С1'
    S2 = 'Судопропускное сооружение С2'
    V1 = 'Водопропускное сооружение В1'
    V2 = 'Водопропускное сооружение В2'
    V3 = 'Водопропускное сооружение В3'
    V4 = 'Водопропускное сооружение В4'
    V5 = 'Водопропускное сооружение В5'
    V6 = 'Водопропускное сооружение В6'
    ZU = 'Здание управления КЗС'
    PS360 = 'ПС 110/35/6кВ №360'
    # PS360 = 'Горская'
    PS86 = 'ПС 110/10/6кВ №86'
    PS223 = 'Бронка'
    KOTLIN = 'Котлин'


class Job:
    def __init__(self):
        self.date: Optional[datetime.date] = None
        self.object: Optional[Objects] = None
        self.system: Optional[Systems] = None
        self.work_type: Optional[str] = None
        self.place: Optional[str] = None
        from duty_schedule import Worker
        self.performer: Optional[Worker] = None
        self.tech_map: Optional[str] = None
        self.equip_name: Optional[str] = None

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
               f'performer:{self._print_str(self.performer)}'

    def __repr__(self):
        return f'obj:{self.object}; place:{self.place}; work:{self.work_type}; ' \
               f'date:{self.date};  sys:{self.system}; performer:{self.performer}'


def find_num_in_str(string: str) -> Optional[int]:
    result = re.findall(r'\d+', string)
    if len(result) > 0:
        return int(result[0])
    else:
        return None


def extract_month(raw_date: str):
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
    for m_name, m_num in month_names.items():
        if m_name in raw_date.lower():
            return m_num
    else:
        raise Exception(f'Cant find month in str \'{raw_date}\'')


def extract_year(raw_date: str):
    raw_year = re.findall(r'\d+', raw_date)
    if len(raw_year) > 0:
        year = int(raw_year[0])
    else:
        raise Exception(f'Cant find year in str \'{raw_date}\'')
    return year


def extract_month_and_year(raw_date: str):
    month = extract_month(raw_date)
    year = extract_year(raw_date)
    # print(raw_date, month, year)  # debug
    return month, year


def extract_place_and_object(raw_place: str):
    places_names = {('ЗУ КЗС',): ('Здание управления КЗС', Objects.ZU),
                    ('Здание управления',): ('Здание управления КЗС', Objects.ZU),
                    ('АМ',): ('С2 АМ', Objects.S2),
                    ('мост',): ('С2 АМ', Objects.S2),
                    ('Бронка',): ('Бронка', Objects.S1),
                    ('ПС', '223'): ('ПС 223', Objects.PS223),
                    ('ПС', '360'): ('ПС 110/35/6кВ №360', Objects.PS360),
                    ('Горская',): ('Горская', Objects.S1),
                    ('ПС', '86'): ('ПС 110/10/6кВ №86', Objects.S1),
                    ('Котлин',): ('ПС Котлин', Objects.S1),
                    ('ПС', 'С1', '110'): ('С1 ПС 110/10кВ', Objects.S1),
                    ('ПС', 'С2', '110'): ('С2 ПС 110/10кВ', Objects.S2),
                    ('Южная', 'С1'): ('С1 Юг', Objects.S1),
                    ('Северная', 'С1'): ('С1 Север', Objects.S1),
                    ('Южная', 'С2'): ('С2 Юг', Objects.S2),
                    ('Северная', 'С2'): ('С2 Север', Objects.S2)
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
        for keyword in i_template:
            if keyword in raw_place:
                continue
            else:
                break
        else:
            return i_place

    # find В1..В6 objects
    search_obj = re.search(r'В\W{,3}(\d)', raw_place)
    if search_obj:
        for obj in Objects:
            if 'В' + search_obj.group(1) in obj.value:
                return obj.value, obj

    # find С1, С2 objects
    search_obj = re.search(r'(С\d)(.*)', raw_place)
    if search_obj:
        for obj in Objects:
            if search_obj.group(1) in obj.value:
                return ''.join(search_obj.groups()), obj

    # raise Exception(f'extract_place_and_object: {raw_place} - нет совпадений с шаблоном')
    return raw_place, 'unknown'


def find_system_by_sheet(sheet_name):
    sheet_names = {'АСУ ТП': Systems.ASU_TP,
                   'АСУ И': Systems.ASU_I,
                   'МОСТ': Systems.ASU_AM,
                   'АМ': Systems.ASU_AM,
                   'ЛВС': Systems.LVS}

    for name, system in sheet_names.items():
        if name in sheet_name:
            return system
    return None


def filter_work_type(work: str):
    work = work.strip(' ,.\t\n')
    work = work.replace('E', 'Е')  # Eng to Rus
    work = work.replace('T', 'Т')  # Eng to Rus
    work = work.replace('O', 'О')  # Eng to Rus
    work = work.replace('ТОЗ', 'ТО3')  # Letter to Num
    return work


def parser_to_jobs(parser) -> List[Job]:
    jobs: List[Job] = []
    month, year = extract_month_and_year(parser.month_year)

    # #TODO это временный код, его необходимо удалить!
    # year = 2021

    # system = find_system_by_sheet(parser.sheet.title)
    for raw_job in parser.raw_data:

        # # TODO это временный код, его необходимо удалить!
        # if month == 2 and raw_job.day == 29:
        #     continue

        job = Job()
        job.place, job.object = extract_place_and_object(raw_job.place)
        job.work_type = filter_work_type(raw_job.work_type)
        job.date = datetime.date(year, month, raw_job.day)
        job.system = parser.system
        job.tech_map = raw_job.tech_map
        job.equip_name = raw_job.equip_name

        from duty_schedule import team_s1, team_s2, team_v, team_tk, team_vols, team_askue

        if job.system in (Systems.LVS, Systems.VOLS):
            team = team_vols
        elif job.system == Systems.TK:
            team = team_tk
        elif job.system in (Systems.ASKUE, Systems.TECH_REG):
            team = team_askue
        elif job.object == Objects.S1:
            team = team_s1
        elif job.object in (Objects.S2, Objects.ZU):
            team = team_s2
        else:
            team = team_v

        from application import duty_schedules

        for schedule in duty_schedules:
            if (schedule.month, schedule.year) == (month, year):
                job.performer = schedule.get_performer(team, job.date)

        jobs.append(job)
    return jobs
