import openpyxl
from openpyxl.styles import NamedStyle, Font, Border, Side, Alignment
from openpyxl.worksheet import Worksheet

from Parser import TableArea


# ws = openpyxl.load_workbook(r'.\input data\11.12.18.xlsx')['Лист1']
# col_dim = ws.column_dimensions
# first_col = col_dim['A']
# col_width = first_col.width
# print(ws.column_dimensions['A'].width)
def style_range(ws, cell_range, border=Border(), fill=None, font=None, alignment=None):
    """
    Apply styles to a range of cells as if they were a single cell.

    :param ws:  Excel worksheet instance
    :param range: An excel range to style (e.g. A1:F20)
    :param border: An openpyxl Border
    :param fill: An openpyxl PatternFill or GradientFill
    :param font: An openpyxl Font object
    """

    top = Border(top=border.top)
    left = Border(left=border.left)
    right = Border(right=border.right)
    bottom = Border(bottom=border.bottom)

    first_cell = ws[cell_range.split(":")[0]]
    if alignment:
        ws.merge_cells(cell_range)
        first_cell.alignment = alignment

    rows = ws[cell_range]
    if font:
        first_cell.font = font

    for cell in rows[0]:
        cell.border = cell.border + top
    for cell in rows[-1]:
        cell.border = cell.border + bottom

    for row in rows:
        l = row[0]
        r = row[-1]
        l.border = l.border + left
        r.border = r.border + right
        if fill:
            for c in row:
                c.fill = fill


def apply_style(sheet: Worksheet, style: NamedStyle, table_area: TableArea):
    for row in sheet.iter_rows(min_row=table_area.first_row, max_row=table_area.last_row,
                               min_col=table_area.first_col, max_col=table_area.last_col):
        for cell in row:
            cell.style = style


# test_out_wb = openpyxl.Workbook()
test_out_wb = openpyxl.load_workbook(r'.\input data\Template.xlsx')
dest_filename = r'test_wb.xlsx'
ws = test_out_wb.worksheets[0]
c1 = ws.cell(2, 2)
c2 = ws.cell(4, 2)
cell_range = ws['B4':'D4']
ws.merge_cells('B4:D4')

style1 = NamedStyle(name='Style1')
bd = Side(style='thin', color='000000')
style1.border = Border(left=bd, top=bd, right=bd, bottom=bd)
style1.alignment = Alignment(vertical='center', horizontal='center')
#
test_out_wb.add_named_style(style1)
# c1.style = style1
# c2.style = style1

c2.value = 'test text'

for col in range(2, 5):
    ws.cell(4, col).style = style1
#
# style_range(ws, 'B4:D4', border=style1.border, alignment=style1.alignment)
# style_range(ws, 'B4:D4', border=style1.border)

test_out_wb.save(filename=dest_filename)
