from collections import namedtuple

from openpyxl.worksheet import Worksheet

from parser import xint

from typing import Tuple, List

# worker = namedtuple('worker', 'full_name phone_number')


class Worker:
    def __init__(self, last_name: str, first_name: str, patronymic: str, phone: str):
        self.last_name = last_name
        self.first_name = first_name
        self.patronymic = patronymic
        self.phone = phone

    def __str__(self):
        return f'{self.last_name} {self.first_name} {self.patronymic} {self.phone} '

    def __repr__(self):
        return self.__str__()


class Team:
    def __init__(self, workers: List[Worker] = None):
        if workers is None:
            self.workers = []
        else:
            self.workers = workers

    def __str__(self):
        return ', '.join(w.last_name for w in self.workers)

    def __repr__(self):
        return self.__str__()

    def get_by_last_name(self, last_name: str) -> Worker:
        for w in self.workers:
            if w.last_name == last_name:
                return w

    def get_by_last_names(self, last_names: Tuple[str]) -> List[Worker]:
        return [self.get_by_last_name(name) for name in last_names]


# workers directory
all_workers = Team([
    Worker('Харитонов', 'Виктор', 'Яковлевич', '+79319642368'),
    Worker('Гусев', 'Михаил', 'Владимирович', '+79675904368'),
    Worker('Мулин', 'Николай', 'Николаевич', '+79112283889'),
    Worker('Кушмылев', 'Евгений', 'Павлович', '+79819134737'),
    Worker('Каприца', 'Анатолий', 'Евгеньевич', '+79046315018'),
    Worker('Макаров', 'Виктор', 'Викторович', '+79312089129'),
    Worker('Горнов', 'Александр', 'Серафимович', '+79111314015'),
    Worker('Подольский', 'Андрей', 'Вениаминович', '+79312531066'),
    Worker('Кокоев', 'Михаил', 'Николаевич', '+79216441993'),
    Worker('Санжара', 'Владимир', 'Александрович', '+79819629285'),
    Worker('Ильин', 'Андрей', 'Владимирович', '+79219303652'),
    Worker('Ястребов', 'Алексей', 'Владимирович', '+79313581975'),
    Worker('Огородников', 'Алексей', 'Юрьевич', '+79313196196')
])

# team_s1 = (w for w in all_workers if w.last_name in ('Гусев', 'Харитонов', 'Мулин', 'Кушмылев'))
# team_s2 = (w for w in all_workers if w.last_name in ('Гусев', 'Харитонов', 'Мулин', 'Кушмылев'))
team_s1 = Team(all_workers.get_by_last_names(('Гусев', 'Харитонов', 'Мулин', 'Кушмылев')))


class DutySchedule:
    def __init__(self, worksheet: Worksheet):
        self.worksheet = worksheet
        self._data_first_col = 3
        self._data_last_col = None
        self._find_last_day_col()
        # номера строк вносятся в порядке убывания приоритета сотрудника для выбора его в качестве руководителя работ
        self.group_s1_rows = (8, 9, 10, 11)
        self.group_s2_rows = (14, 13)
        self.group_v_rows = (17, 18, 19)
        self.group_vols_rows = (23, 21, 22)
        self.group_tk_rows = (22, 21, 23)
        self.all_workers_rows = self.group_s1_rows + self.group_s2_rows + self.group_v_rows + \
                                self.group_vols_rows + self.group_tk_rows

    def _find_last_day_col(self):
        days_row = 2
        table_max_col = 50
        for col in range(self._data_first_col, table_max_col + 1):
            cell = self.worksheet.cell(days_row, col).value
            if type(xint(cell)) is not int:
                self._data_last_col = col - 1
                break
        else:
            print('Последняя дата не найдена')

    def _day2col(self, day: int):
        col = day + self._data_first_col - 1
        if col <= self._data_last_col:
            return col
        else:
            return None

    def _is_workday(self, day: int):
        col = self._day2col(day)
        if col is None:
            return None
        else:
            for row in self.all_workers_rows:
                cell = xint(self.worksheet.cell(row, col).value)
                if cell == 8 or cell == 7:
                    return True
            else:
                return False

    # def get_s1_worker(self, day: int):
    #     for row in self.group_s1_rows
    #         if


if __name__ == '__main__':
    print(all_workers)

    print(team_s1)
