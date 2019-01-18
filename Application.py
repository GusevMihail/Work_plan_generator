# from pathlib import Path
import openpyxl

import Parser
import Pre_processing



class Job:
    def __init__(self):
        self.date = None
        self.object = None
        self.system = None
        self.work_type = None
        self.place = None
        self.worker = None


if __name__ == "__main__":
    workbook_asu = openpyxl.load_workbook(r'.\input data\5. Графики на 05.18 АСУ.xlsx')
    for sheet_name in workbook_asu.sheetnames:  # find necessary worksheets by names
        # TODO move system name finding to Pre_processing module
        for name, system in Pre_processing.sheet_names.items():
            if name in sheet_name:
                parse_sheet(wb[sheet_name], system)
                break
