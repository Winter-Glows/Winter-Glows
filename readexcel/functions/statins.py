from datetime import date
from pathlib import Path
from pprint import pprint

import xlrd
from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage


def get_result(record: dict, database, CFG, mode, tpl):

    # record = {'genetype': {"TNB-10n": "TT",
    #           "TTL-1_1": "TC",
    #           "TTL-1_2": "TC"
    #         },}

    advice = {}

    for info in record.keys():
        exec(f"advice['{info}'] = record['{info}']")
        advice[info] = advice[info] if advice[info] else '-'
    for key in advice['genetype'].keys():
        exec(f"advice['{key}'] = advice['genetype']['{key}']")
        # if advice[key] in ('CT','TC'):
        #     advice[key] = 'TC'
        # print(advice[key])

    advice['SLCO1B1'] = advice['TNB-10n']
    advice['apoe_T388C'] = advice['TTL-1_1']
    advice['apoe_C526T'] = advice['TTL-1_2']

    pic_dir = Path(CFG['Dirs']['pictures'] + '/' + 'statins')

    # import xlrd
    # database = '../interpretations/statins.xlsx'
    work = xlrd.open_workbook(database)
    sheet = work.sheet_by_index(0)
    rows = sheet.nrows
    row_values = {}
    for row in range(rows):
        row_values[row] = sheet.row_values(rowx = row)
        if row_values[row][0] == 'SLCO1B1【他汀与肌病风险】':
            if row_values[row][1] == advice['SLCO1B1']:
                advice['result_SLCO1B1'] = row_values[row][2]
                exec(f"iamge_path_SLCO1B1 = pic_dir / 'SLCO1B1_{advice['SLCO1B1']}.png'")
        if row_values[row][0] == 'ApoE【他汀与降脂效果】':
            if row_values[row][1] == advice['apoe_T388C'] and row_values[row][2] == advice['apoe_C526T']:
                advice['result_apoe'] = row_values[row][3]
                exec(f"iamge_path_apoe_T388C = pic_dir / 'apoe_T388C_{advice['apoe_T388C']}.png'")
                exec(f"iamge_path_apoe_C526T = pic_dir / 'apoe_C526T_{advice['apoe_C526T']}.png'")
    # pprint(advice)

    for i in ['SLCO1B1','apoe_T388C','apoe_C526T']:
        exec(f"advice['peak_figure_{i}'] = InlineImage(tpl, str(iamge_path_{i}.absolute()), width=Mm(45), height=Mm(27))")

    # advice['zk_entrust_time'] = record['zk_entrust_time']
    advice['zk_accept_time'] = str(advice['receive_time'][:10]).replace('-','.')
    advice['zk_report_time'] = str(date.today()).replace('-','.')
    
    return advice