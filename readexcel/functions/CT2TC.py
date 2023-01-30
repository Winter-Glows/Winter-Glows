from pprint import pprint

import xlrd

database = '../interpretations/statins.xlsx'
work = xlrd.open_workbook(database)
sheet = work.sheet_by_index(0)
rows = sheet.nrows
row_values = {}
for row in range(rows):
    row_values[row] = sheet.row_values(row)
    if row_values[row][1] in ('CT','TC'):
        row_values[row][1] = 'CT'
    if row_values[row][2] in ('CT','TC'):
        row_values[row][2] = 'CT'
    pprint(row_values[row])