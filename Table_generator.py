import openpyxl
from openpyxl.styles import NamedStyle, Font, Border, Side, Alignment, PatternFill
from openpyxl.worksheet import Worksheet

from Parser import TableArea

from datetime import date

FIRST_COL = 1
LAST_COL = 9
FIRST_DATA_ROW = 10

thin_side = Side(style='thin', color='000000')
thin_border = Border(left=thin_side, top=thin_side, right=thin_side, bottom=thin_side)

current_row = FIRST_DATA_ROW


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


def apply_named_style(worksheet: Worksheet, style: NamedStyle, table_area: TableArea):
    for row in worksheet.iter_rows(min_row=table_area.first_row, max_row=table_area.last_row,
                                   min_col=table_area.first_col, max_col=table_area.last_col):
        for cell in row:
            cell.style = style


def apply_style(worksheet: Worksheet, table_area: TableArea, border: Border = None, alignment: Alignment = None,
                fill: PatternFill = None, font: Font = None):
    for row in worksheet.iter_rows(min_row=table_area.first_row, max_row=table_area.last_row,
                                   min_col=table_area.first_col, max_col=table_area.last_col):
        for cell in row:
            if border is not None:
                cell.border = border
            if alignment is not None:
                cell.alignment = alignment
            if fill is not None:
                cell.fill = fill
            if font is not None:
                cell.font = font


def get_template(path: str):
    wb_template = openpyxl.load_workbook(path)
    ws_template = wb_template.active
    header_1 = TableArea(first_row=1, last_row=7, first_col=FIRST_COL, last_col=LAST_COL)
    header_2 = TableArea(first_row=8, last_row=9, first_col=FIRST_COL, last_col=LAST_COL)
    apply_style(ws_template, header_2, border=thin_border)
    return wb_template


def write_date(ws: Worksheet, table_date: date):
    date_cell = ws.cell(row=6, column=1)
    date_cell.value = table_date


if __name__ == '__main__':
    # test_out_wb = openpyxl.Workbook()
    template_filename = r'.\input data\Template.xlsx'
    test_out_wb = get_template(template_filename)
    test_ws = test_out_wb.worksheets[0]
    test_date = date(2019, 1, 20)
    write_date(test_ws, test_date)

    # style1 = NamedStyle(name='Style1')
    # style1.border = Border(left=thin_side, top=thin_side, right=thin_side, bottom=thin_side)
    # style1.alignment = Alignment(vertical='center', horizontal='center')
    # style1.fill = PatternFill(fill_type='solid', start_color='ff8327', end_color='ff8327')
    # test_out_wb.add_named_style(style1)

    dest_filename = r'test_wb.xlsx'
    test_out_wb.save(filename=dest_filename)
