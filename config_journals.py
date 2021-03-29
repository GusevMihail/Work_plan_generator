from collections import namedtuple
from enum import Enum

from pre_processing import Systems, Objects

JournalGeneratorConfig = namedtuple('journal_generator_config', ('system', 'object', 'place_filter'))


class PFilters(Enum):
    south = 'Юг'
    north = 'Север'


default_header_ASU = (('Дата', 10), ('Место', 30), ('Тип', 7), ('Тех. карта', 50), ('Исполнитель', 15))

batch_ASU_journals = {'АСУ ТП С1 Север': JournalGeneratorConfig(Systems.ASU_TP, Objects.S1, 'Север'),
                      'АСУ ТП С1 Юг': JournalGeneratorConfig(Systems.ASU_TP, Objects.S1, 'Юг'),
                      'АСУ ТП С2 Север': JournalGeneratorConfig(Systems.ASU_TP, Objects.S2, 'Север'),
                      'АСУ ТП С2 Юг': JournalGeneratorConfig(Systems.ASU_TP, Objects.S2, 'Юг'),
                      'АСУ ТП В1': JournalGeneratorConfig(Systems.ASU_TP, Objects.V1, None),
                      'АСУ ТП В2': JournalGeneratorConfig(Systems.ASU_TP, Objects.V2, None),
                      'АСУ ТП В3': JournalGeneratorConfig(Systems.ASU_TP, Objects.V3, None),
                      'АСУ ТП В4': JournalGeneratorConfig(Systems.ASU_TP, Objects.V4, None),
                      'АСУ ТП В5': JournalGeneratorConfig(Systems.ASU_TP, Objects.V5, None),
                      'АСУ ТП В6': JournalGeneratorConfig(Systems.ASU_TP, Objects.V6, None),
                      'АСУ ТП ЗУ': JournalGeneratorConfig(Systems.ASU_TP, Objects.ZU, None),

                      'АСУ АМ С2': JournalGeneratorConfig(Systems.ASU_AM, Objects.S2, None),

                      'АСУ И С1 Север': JournalGeneratorConfig(Systems.ASU_I, Objects.S1, 'Север'),
                      'АСУ И С1 Юг': JournalGeneratorConfig(Systems.ASU_I, Objects.S1, 'Юг'),
                      'АСУ И С2 Север': JournalGeneratorConfig(Systems.ASU_I, Objects.S2, 'Север'),
                      'АСУ И С2 Юг': JournalGeneratorConfig(Systems.ASU_I, Objects.S2, 'Юг'),
                      'АСУ И ЗУ': JournalGeneratorConfig(Systems.ASU_I, Objects.ZU, None),

                      'ВОЛС': JournalGeneratorConfig(Systems.VOLS, None, None),
                      'ТК М2': JournalGeneratorConfig(Systems.TK, None, None),
                      }

default_header_ASKUE = (
    ('Дата', 10), ('Место', 30), ('Оборудование', 30), ('Тип', 7), ('Тех. карта', 32), ('Исполнитель', 15))

batch_ASKUE_journals = {'АИИСКУЕ': JournalGeneratorConfig(Systems.ASKUE, None, None),
                        'Тех Учет': JournalGeneratorConfig(Systems.TECH_REG, None, None),
                        'ЛВС': JournalGeneratorConfig(Systems.LVS, None, None)
                        }

batch_ASU_test = {'АСУ ТП С1 Север': JournalGeneratorConfig(Systems.ASU_TP, Objects.S1, 'Север'),
                  'АСУ И С1 Север': JournalGeneratorConfig(Systems.ASU_I, Objects.S1, 'Север')}
