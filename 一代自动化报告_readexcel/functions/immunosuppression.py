from datetime import date
from pathlib import Path

import xlrd
from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage


def get_result(record: dict, database, CFG, mode , tpl):

    # record = {'genetype':{'ABCB1_rs1045642':'CC', 
    #         'UGT1A8':'CT', 
    #         'DROSHA':'GA', 
    #         'CYP1A2':'CC', 
    #         'ESR1':'TT', 
    #         'POR':'CC', 
    #         'TPMT':'GG', 
    #         'NUDT15':'CT', 
    #         'CYP3A5_3':'AA', 
    #         'ABCB1_rs2032582':'AA', 
    #         'MTHFR_1':'CT', 
    #         'TCF7L2':'TT'
    #         },
    #         'flag_jjpnsl':True,
    #         "flag_lcpl":True,
    #         "flag_mfsz":True,
    #         "flag_qds":False,
    #         "flag_lfmt":False,
    #         "flag_hbms":False,
    #         "flag_tkms":False,
    #         "flag_hlxa":False,}

    advice = {}
    for info in record.keys():
        exec(f"advice['{info}'] = record['{info}']")
        advice[info] = advice[info] if advice[info] else '-'
    for key in advice['genetype'].keys():
        exec(f"advice['{key}'] = advice['genetype']['{key}']")

    # import xlrd
    # database = '../interpretations/immunosuppression.xlsx'
    work = xlrd.open_workbook(database)
    sheet = work.sheet_by_index(0)
    rows = sheet.nrows
    row_values = {}
    for row in range(rows):
        row_values[row] = sheet.row_values(rowx = row)
        if advice['flag_jjpnsl'] == True:
            if row_values[row][0] == '甲基泼尼松龙':
                if row_values[row][1] == advice['ABCB1_rs1045642']:
                    advice['tip_Methylprednisolone'] = row_values[row][2]
        if advice['flag_lcpl'] == True:
            if row_values[row][0] == '硫唑嘌呤':
                if row_values[row][1] == advice['TPMT']:
                    advice['tip_azathioprine_TPMT'] = row_values[row][2]
                if row_values[row][3] == advice['NUDT15']:
                    advice['tip_azathioprine_NUDT15'] = row_values[row][4]
        if advice['flag_mfsz'] == True:
            if row_values[row][0] == '霉酚酸酯':
                if row_values[row][1] == advice['UGT1A8']:
                    advice['tip_mycophenolatemofetil_UGT1A8'] = row_values[row][2]
                if row_values[row][3] == advice['ABCB1_rs2032582']:
                    advice['tip_mycophenolatemofetil_ABCB1'] = row_values[row][4]
        if advice['flag_qds'] == True:
            if row_values[row][0] == '强的松':
                if row_values[row][1] == advice['DROSHA']:
                    advice['tip_prednisone_DROSHA'] = row_values[row][2]
                if row_values[row][3] == advice['ABCB1_rs1045642']:
                    advice['tip_prednisone_ABCB1'] = row_values[row][4]
        if advice['flag_lfmt'] == True:
            if row_values[row][0] == '来氟米特':
                if row_values[row][1] == advice['CYP1A2']:
                    advice['tip_Leflunomide_CYP1A2'] = row_values[row][2]
                if row_values[row][3] == advice['ESR1']:
                    advice['tip_Leflunomide_ESR1'] = row_values[row][4]
        if advice['flag_hbms'] == True:
            if row_values[row][0] == '环孢霉素':
                if row_values[row][1] == advice['CYP3A5_3']:
                    advice['tip_cyclosporine_CYP3A5_3'] = row_values[row][2]
                if row_values[row][3] == advice['TCF7L2']:
                    advice['tip_cyclosporine_TCF7L2'] = row_values[row][4]
                if row_values[row][5] == advice['ABCB1_rs2032582']:
                    advice['tip_cyclosporine_ABCB1'] = row_values[row][6]
        if advice['flag_tkms'] == True:
            if row_values[row][0] == '他克莫司':
                if row_values[row][1] == advice['CYP3A5_3']:
                    advice['tip_Tacrolimus_CYP3A5_3'] = row_values[row][2]
                if row_values[row][3] == advice['POR']:
                    advice['tip_Tacrolimus_POR'] = row_values[row][4]
        if advice['flag_hlxa'] == True:
            if row_values[row][0] == '环磷酰胺':
                if row_values[row][1] == advice['MTHFR_1']:
                    advice['tip_cyclophosphamide_MTHFR_1'] = row_values[row][2]
    # print(advice)

    drugs = []
    drugs_dict = {'hbms':'环孢霉素','tkms':'他克莫司','hlxa':'环磷酰胺','jjpnsl':'甲基泼尼松龙',
                    'lcpl':'硫唑嘌呤','mfsz':'霉酚酸酯','qds':'强的松','lfmt':'来氟米特'}
    for i,j in drugs_dict.items():
        if advice[f'flag_{i}'] == True:
            drugs.append(j)
    drugs = '、'.join(drugs)

    advice['drugs'] = drugs
    # advice['zk_entrust_time'] = record['zk_entrust_time']
    advice['zk_accept_time'] = str(advice['receive_time'][:10]).replace('-','.')
    advice['zk_report_time'] = str(date.today()).replace('-','.')

    return advice