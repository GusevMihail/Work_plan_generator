import openpyxl

import duty_schedule
from datetime import date

if __name__ == '__main__':
    f = open(r'duty_list.txt', 'w')
    workbook = openpyxl.load_workbook(r'.\input data\Test график дежурств 19.02.xlsx')
    schedule = duty_schedule.DutySchedule(workbook.worksheets[0], duty_schedule.all_workers)
    teams = (duty_schedule.team_s1, duty_schedule.team_s2, duty_schedule.team_v, duty_schedule.team_tk,
             duty_schedule.team_vols, duty_schedule.team_askue)
    for day in range(1, 29):
        for team in teams:
            performer = schedule.get_performer(team, date(2019, 2, day)).last_name
            f.write(f'{str(day).ljust(3)}{str(team.name).ljust(15)}{performer}\n')
    f.close()
