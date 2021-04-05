from collections import namedtuple

from pre_processing import Objects, Systems

CalendarDescription = namedtuple('CalendarDescription', ['name', 'objects', 'systems'])

calendars_settings = [
    CalendarDescription(name='В1-В6', objects=[Objects.V1, Objects.V2, Objects.V3, Objects.V4, Objects.V5, Objects.V6],
                        systems=[Systems.ASU_TP, ]),

    CalendarDescription(name='C1', objects=[Objects.S1, ],
                        systems=[Systems.ASU_TP, Systems.ASU_I]),

    CalendarDescription(name='C2', objects=[Objects.S2, Objects.ZU],
                        systems=[Systems.ASU_TP, Systems.ASU_AM, Systems.ASU_I]),

    CalendarDescription(name='Энергетика', objects=None,
                        systems=[Systems.ASKUE, Systems.TECH_REG, Systems.TK]),

    CalendarDescription(name='Оптика', objects=None,
                        systems=[Systems.LVS, Systems.VOLS])
]
