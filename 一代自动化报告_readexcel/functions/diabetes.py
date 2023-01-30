from datetime import date
from pathlib import Path

import xlrd
from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage


def get_result(record: dict, database, CFG, mode, tpl):  #

    # record = {"genetype": {
    #   "ABCC8": "TG",
    #   "CYP2C9": "AA",
    #   "GLP1R_rs10305420": "CC",
    #   "GLP1R_rs6923761": "AG",
    #   "PPARG": "CC",
    #   "SLC22A2": "GG",
    #   "SLCO1B1": "TT"
    # },}

    advice = {}

    for info in record.keys():
        exec(f"advice['{info}'] = record['{info}']")
        advice[info] = advice[info] if advice[info] else '-'

    for key in advice['genetype'].keys():
        exec(f"advice['{key}'] = advice['genetype']['{key}']")

    # import xlrd
    # database = '../interpretations/diabetes.xlsx'
    work = xlrd.open_workbook(database)
    sheet = work.sheet_by_index(0)
    rows = sheet.nrows
    row_values = {}
    for row in range(rows):
        row_values[row] = sheet.row_values(rowx = row)
        if row_values[row][0] == '磺脲类':
            if row_values[row][1] == advice['CYP2C9'] and row_values[row][2] == advice['ABCC8']:
                advice['text_1'] = row_values[row][3]
                advice['tip_1'] = row_values[row][4]
        elif row_values[row][0] == '双胍类':
            if row_values[row][1] == advice['SLC22A2']:
                advice['text_2'] = row_values[row][3]
                advice['tip_2'] = row_values[row][4]
        elif row_values[row][0] == '噻唑烷二酮类':
            if row_values[row][1] == advice['PPARG']:
                advice['text_3'] = row_values[row][3]
                advice['tip_3'] = row_values[row][4]
        elif row_values[row][0] == '氯茴苯酸类':
            if row_values[row][1] == advice['CYP2C9'] and row_values[row][2] == advice['SLCO1B1']:
                advice['text_4'] = row_values[row][3]
                advice['tip_4'] = row_values[row][4]
        elif row_values[row][0] == 'DPP-4 抑制剂':
            if row_values[row][1] == advice['GLP1R_rs6923761']:
                advice['text_5'] = row_values[row][3]
                advice['tip_5'] = row_values[row][4]
        elif row_values[row][0] == 'GLP-1 受体激动剂':
            if row_values[row][1] == advice['GLP1R_rs10305420']:
                advice['text_6'] = row_values[row][3]
                advice['tip_6'] = row_values[row][4]
    # print(advice)

    # advice['zk_entrust_time'] = record['zk_entrust_time']
    advice['zk_accept_time'] = str(advice['receive_time'][:10]).replace('-','.')
    advice['zk_report_time'] = str(date.today()).replace('-','.')

    return advice
    