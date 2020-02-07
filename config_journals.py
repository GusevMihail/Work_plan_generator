from collections import namedtuple
from enum import Enum

from pre_processing import Systems, Objects

JournalGeneratorConfig = namedtuple('journal_generator_config', ('system', 'object', 'place_filter'))


class PFilters(Enum):
    south = 'Юг'
    north = 'Север'


batch_ASU_journals = {'АСУ ТП С1 Север': JournalGeneratorConfig(Systems.ASU_TP, Objects.S1, PFilters.north),
                      'АСУ ТП С1 Юг': JournalGeneratorConfig(Systems.ASU_TP, Objects.S1, PFilters.north),
                      'АСУ ТП С2 Север': JournalGeneratorConfig(Systems.ASU_TP, Objects.S2, PFilters.north),
                      'АСУ ТП С2 Юг': JournalGeneratorConfig(Systems.ASU_TP, Objects.S2, PFilters.north),
                      # 'АСУ ТП С2 АМ':       JournalGeneratorConfig(Systems.ASU_TP, Objects.S2, PFilters.north),
                      'АСУ ТП В1': JournalGeneratorConfig(Systems.ASU_TP, Objects.V1, None),
                      'АСУ ТП В2': JournalGeneratorConfig(Systems.ASU_TP, Objects.V2, None),
                      'АСУ ТП В3': JournalGeneratorConfig(Systems.ASU_TP, Objects.V3, None),
                      'АСУ ТП В4': JournalGeneratorConfig(Systems.ASU_TP, Objects.V4, None),
                      'АСУ ТП В5': JournalGeneratorConfig(Systems.ASU_TP, Objects.V5, None),
                      'АСУ ТП В6': JournalGeneratorConfig(Systems.ASU_TP, Objects.V6, None),
                      'АСУ ТП ЗУ': JournalGeneratorConfig(Systems.ASU_TP, Objects.ZU, None),

                      'АСУ И С1 Север': JournalGeneratorConfig(Systems.ASU_I, Objects.S1, PFilters.north),
                      'АСУ И С1 Юг': JournalGeneratorConfig(Systems.ASU_I, Objects.S1, PFilters.north),
                      'АСУ И С2 Север': JournalGeneratorConfig(Systems.ASU_I, Objects.S2, PFilters.north),
                      'АСУ И С2 Юг': JournalGeneratorConfig(Systems.ASU_I, Objects.S2, PFilters.north),
                      # 'АСУ ТП С2 АМ':       JournalGeneratorConfig(Systems.ASU_TP, Objects.S2, PFilters.north),
                      'АСУ И ЗУ': JournalGeneratorConfig(Systems.ASU_I, Objects.ZU, None)

                      # 'ЛВС':                JournalGeneratorConfig(Systems.LVS, None, None)

                      }
