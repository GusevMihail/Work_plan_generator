import unittest
from pathlib import Path
import Parser
import openpyxl

class TestParser(unittest.TestCase):
    path_asu_real = Path(
        r"c:\Users\Mihail\PycharmProjects\Work_plan_generator\input data\\5. Графики на 05.18 АСУ.xlsx")
    wb_asu_real = openpyxl.load_workbook(str(path_asu_real))
    sheet_asu_109 = wb_asu_real['109. МОСТ']

    def test_find_data_boundaries(self):
        self.assertEqual(Parser.ParserAsu(sheet_asu_109),)
