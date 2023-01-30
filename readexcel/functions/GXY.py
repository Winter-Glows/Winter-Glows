from datetime import date
from pathlib import Path

import xlrd
from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage


def get_result(record: dict, database, CFG, mode, tpl):

    # record = {"genetype": {
    #     "ADD1": "GT",
    #     "ADRB1": "GC",
    #     "AGTR1": "AA",
    #     "CYP2C9": "AC",
    #     "CYP2D6": "CC",
    #     "CYP3A5": "AG",
    #     "NEDD4L": "AG"
    # },}

    advice = {}
    record['genetype']['ACE'] = 'DD'

    for info in record.keys():
        exec(f"advice['{info}'] = record['{info}']")
        advice[info] = advice[info] if advice[info] else '-'

    for key in advice['genetype'].keys():
        exec(f"advice['{key}'] = advice['genetype']['{key}']")

    # import xlrd
    # database = '../interpretations/GXY.xlsx'
    work = xlrd.open_workbook(database)
    sheet = work.sheet_by_index(0)
    rows = sheet.nrows
    row_values = {}
    for row in range(rows):
        row_values[row] = sheet.row_values(rowx = row)
        if row_values[row][0] == '血管紧张素Ⅱ受体拮抗剂':
            if row_values[row][1] == advice['CYP2C9'] and row_values[row][2] == advice['AGTR1']:
                advice['text_1'] = row_values[row][3]
                advice['tip_1'] = row_values[row][4]
        if row_values[row][0] == '血管紧张素转换酶抑制剂':
            if row_values[row][1] == advice['ACE']:
                advice['text_2'] = row_values[row][3]
                advice['tip_2'] = row_values[row][4]
        if row_values[row][0] == 'β受体阻断药':
            if row_values[row][1] == advice['CYP2D6'] and row_values[row][2] == advice['ADRB1']:
                advice['text_3'] = row_values[row][3]
                advice['tip_3'] = row_values[row][4]
        if row_values[row][0] == '钙拮抗剂':
            if row_values[row][1] == advice['CYP3A5']:
                advice['text_4'] = row_values[row][3]
                advice['tip_4'] = row_values[row][4]
        if row_values[row][0] == '利尿药':
            if row_values[row][1] == advice['ADD1'] and row_values[row][2] == advice['NEDD4L']:
                advice['text_5'] = row_values[row][3]
                advice['tip_5'] = row_values[row][4]
    # print(advice)

    # advice['zk_entrust_time'] = record['zk_entrust_time']
    advice['zk_accept_time'] = str(advice['receive_time'][:10]).replace('-','.')
    advice['zk_report_time'] = str(date.today()).replace('-','.')
    
    return advice