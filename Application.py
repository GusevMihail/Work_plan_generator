# from pathlib import Path
import openpyxl

import Parser
import Pre_processing

if __name__ == "__main__":
    workbook_asu = openpyxl.load_workbook(r'.\input data\5. Графики на 05.18 АСУ.xlsx')
    for sheet_name in workbook_asu.sheetnames:  # find necessary worksheets by names
        system = Pre_processing.extract_system(sheet_name)
        if system is not None:
            parser = Parser.ParserAsu(workbook_asu[sheet_name])

    print(workbook_asu.sheetnames)
