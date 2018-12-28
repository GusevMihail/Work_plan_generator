from pathlib import Path

from Pre_processing import extract_place


class Job:
    def __init__(self):
        self.date = None
        self.object = None
        self.system = None
        self.work_type = None
        self.place = None
        self.worker = None


if __name__ == "__main__":
    jobs_schedule_asu = Path(
        r"c:\Users\Mihail\PycharmProjects\Work_plan_generator\input data\\5. Графики на 05.18 АСУ.xlsx")
    # parser_asu(jobs_schedule_asu)

    test_raw_places = open('test raw places.txt')
    for line in test_raw_places:
        print(f'{line}  -->>  { extract_place(line)}')
    # print(extract_place('Судопропускное сооружение Са1 Юг ДКФ'))
