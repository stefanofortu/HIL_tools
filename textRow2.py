import xlrd
import xlwt
from xlutils.copy import copy

col = 0
row = 0
rb = xlrd.open_workbook('test.xls', formatting_info=True)
r_sheet = rb.sheet_by_index(0)
text_cell = r_sheet.cell_value(row, col)
print(text_cell)
book = copy(rb)
first_sheet = book.get_sheet(0)
font1 = xlwt.easyfont('struck_out true, color_index red')
font2 = xlwt.easyfont('color_index green')
seg1 = (text_cell[0:10], font1)
seg2 = (text_cell[10:], font2)
first_sheet.write_rich_text(row, col, [seg1, seg2])
book.save('test2.xls')