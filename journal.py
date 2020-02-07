from abc import ABCMeta, abstractmethod
from typing import List, Union

import pandas as pd

from pre_processing import Job, Systems, Objects


def jobs_to_df(jobs: List[Job]) -> pd.DataFrame:
    columns = ('date', 'system', 'object', 'place', 'work_type', 'tech_map', 'equip_name', 'performer',)
    result = pd.DataFrame(columns=columns,
                          data=((j.date, j.system, j.object, j.place, j.work_type,
                                 j.tech_map, j.equip_name, j.performer)
                                for j in jobs))
    return result


class Journal(metaclass=ABCMeta):
    def __init__(self, jobs_df: pd.DataFrame):
        self.df = jobs_df
        self.journal: Union[None, pd.DataFrame] = None

    @abstractmethod
    def make_journal(self, sys: Systems, obj: Objects, place: str):
        raise NotImplementedError

    def save_journal(self, file_name='default', folder=r'./output data/journals/'):
        if self.journal is None:
            raise ValueError('make journal before save it!')
        else:
            full_file_name = folder + self.df.date[0].strftime('%y %m ') + file_name + '.xlsx'
            sheet_name = 'Sheet1'
            with pd.ExcelWriter(path=full_file_name, date_format='DD.MM.YY') as writer:
                header = ('Дата', 'Место', 'Тип', 'Тех. карта', 'Исполнитель')
                self.journal.to_excel(writer, sheet_name, header=header, index=False)
                self._set_columns_width(writer.sheets[sheet_name])

    @staticmethod
    def _set_columns_width(worksheet):
        from openpyxl.utils import get_column_letter
        col_widths = (10, 15, 7, 50, 15)
        for i, width in enumerate(col_widths):
            worksheet.column_dimensions[get_column_letter(i + 1)].width = width


class JournalASU(Journal):
    def make_journal(self, sys: Systems, obj: Objects, place_filter: str = None):
        df = self.df  # сокращение записи
        journal = df[df.system == sys] \
            [df.object == obj] \
            .drop(['system', 'object'], axis=1) \
            .reindex(columns=['date', 'place', 'work_type', 'tech_map', 'performer']) \
            .sort_values(by=['date', 'place', 'work_type'])

        if place_filter:
            journal = journal[journal.place.str.contains(place_filter)]

        # journal.date = journal.date.apply(lambda d: d.strftime('%d.%m.%y'))
        journal.performer = journal.performer.apply(lambda w: w.last_name)

        self.journal = journal
