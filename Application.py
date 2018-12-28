from pathlib import Path

import re


class Job:
    def __init__(self):
        self.date = None
        self.object = None
        self.system = None
        self.work_type = None
        self.place = None
        self.worker = None


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


def extract_place(raw_place: str):
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


if __name__ == "__main__":
    jobs_schedule_asu = Path(
        r"c:\Users\Mihail\PycharmProjects\Work_plan_generator\input data\\5. Графики на 05.18 АСУ.xlsx")
    # parser_asu(jobs_schedule_asu)

    test_raw_places = open('test raw places.txt')
    for line in test_raw_places:
        print(f'{line}  -->>  {extract_place(line)}')
    # print(extract_place('Судопропускное сооружение Са1 Юг ДКФ'))
