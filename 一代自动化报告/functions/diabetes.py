from datetime import date
from pathlib import Path

from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage


def get_result(record: dict, CFG, mode, tpl):  #

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

    gene_type_CYP2C9_A1075C = advice['CYP2C9']
    gene_type_ABCC8_G4105GT = advice['ABCC8']
    gene_type_SLC22A2_T808G = advice['SLC22A2']
    gene_type_PPARG_C34G = advice['PPARG']
    gene_type_SLCO1B1_T521C = advice['SLCO1B1']
    gene_type_GLP1R_G502A = advice['GLP1R_rs6923761']
    gene_type_GLP1R_C20T = advice['GLP1R_rs10305420']

    # 1. 磺脲类    4. 氯茴苯酸类
    if gene_type_CYP2C9_A1075C == 'AA':
        if gene_type_ABCC8_G4105GT in ('GG','TT'):
            advice['text_1'] = f'磺脲类： CYP2C9基因型为AA，药物敏感性较好，建议使用该类药物；ABCC8基因型为{gene_type_ABCC8_G4105GT}，药物敏感性较好，建议使用该类药物。'
            advice['tip_1'] = '推荐'
        elif gene_type_ABCC8_G4105GT in ('GT','TG'):
            advice['text_1'] = f'磺脲类： CYP2C9基因型为AA，药物敏感性较好，建议使用该类药物；ABCC8基因型为{gene_type_ABCC8_G4105GT}，药物敏感性较差，建议适当增加剂量或换药。'
            advice['tip_1'] = '可用（↑）'
        if gene_type_SLCO1B1_T521C == 'TT':
            advice['text_4'] = '氯茴苯酸类：CYP2C9基因型为AA，SLCO1B1基因型为TT，药物清除率正常，建议使用常规剂量。'
            advice['tip_4'] = '推荐'
        elif gene_type_SLCO1B1_T521C in ('TC','CT','CC'):
            advice['text_4'] = f'氯茴苯酸类：CYP2C9基因型为AA，SLCO1B1基因型为{gene_type_SLCO1B1_T521C}，药物清除率较低，建议适当降低剂量。'
            advice['tip_4'] = '可用（↓）'
    if gene_type_CYP2C9_A1075C in ('AC','CA','CC'):
        if gene_type_ABCC8_G4105GT in ('GG','TT'):
            advice['text_1'] = f'磺脲类：CYP2C9基因型为{gene_type_CYP2C9_A1075C}，药物敏感性较好，药效好，但低血糖的风险高，建议适当降低剂量或换药；ABCC8基因型为{gene_type_ABCC8_G4105GT}，药物敏感性较好，建议使用该类药物。'
            advice['tip_1'] = '可用（↓）'
        elif gene_type_ABCC8_G4105GT in ('GT','TG'):
            advice['text_1'] = f'磺脲类：CYP2C9基因型为{gene_type_CYP2C9_A1075C}，药物敏感性较好，药效好，但低血糖的风险高，建议适当降低剂量或换药；ABCC8基因型为{gene_type_ABCC8_G4105GT}，药物敏感性较差，建议适当增加剂量或换药。'
            advice['tip_1'] = '慎用'
        advice['text_4'] = f'氯茴苯酸类：CYP2C9基因型为{gene_type_CYP2C9_A1075C}，SLCO1B1基因型为{gene_type_SLCO1B1_T521C}，药物清除率较低，建议适当降低剂量。'
        advice['tip_4'] = '可用（↓）'
    
    # 2. 双胍类
    if gene_type_SLC22A2_T808G == 'GG':
        advice['text_2'] = '双胍类：SLC22A2基因型为GG，转运功能正常，二甲双胍清除率正常，建议使用常规剂量。'
        advice['tip_2'] = '推荐'
    elif gene_type_SLC22A2_T808G in ('GT','TG','TT'):
        advice['text_2'] = f'双胍类：SLC22A2基因型为{gene_type_SLC22A2_T808G}，转运功能降低，致使肾脏对二甲双胍的清除率减慢，体内药物浓度升高，降糖效应增强，发生低血糖风险增加，建议降低用药剂量。'
        advice['tip_2'] = '可用（↓）'

    # 3. 噻唑烷二酮类
    if gene_type_PPARG_C34G == 'CC':
        advice['text_3'] = '噻唑烷二酮类：PPARG基因型为CC，药物敏感性较差，建议适当增加剂量或换药。'
        advice['tip_3'] = '可用（↑）'
    elif gene_type_PPARG_C34G in ('CG','GC','GG'):
        advice['text_3'] = f'噻唑烷二酮类：PPARG基因型为{gene_type_PPARG_C34G}，药物敏感性较好，建议使用该类药物。'
        advice['tip_3'] = '推荐'
    
    # 5. DPP-4 抑制剂
    if gene_type_GLP1R_G502A in ('GA','AG','GG'):
        advice['text_5'] = f'DPP-4 抑制剂：GLP1R基因型为{gene_type_GLP1R_G502A}，药物敏感性较好，建议使用该类药物。'
        advice['tip_5'] = '推荐'
    elif gene_type_GLP1R_G502A == 'AA':
        advice['text_5'] = 'DPP-4 抑制剂：GLP1R基因型为AA，药物敏感性较差，建议适当增加剂量或换药。'
        advice['tip_5'] = '可用（↑）'
    
    # 6. GLP-1 受体激动剂
    if gene_type_GLP1R_C20T == 'CC':
        advice['text_6'] = 'GLP-1 受体激动剂：GLP1R基因型为CC，药物敏感性较好，建议使用该类药物。'
        advice['tip_6'] = '推荐'
    elif gene_type_GLP1R_C20T in ('CT','TC','TT'):
        advice['text_6'] = f'GLP-1 受体激动剂：GLP1R基因型为{gene_type_GLP1R_C20T}，药物敏感性较差，建议适当增加剂量或换药。'
        advice['tip_6'] = '可用（↑）'
    # print(advice)

    # advice['zk_entrust_time'] = record['zk_entrust_time']
    advice['zk_accept_time'] = str(advice['receive_time'][:10]).replace('-','.')
    advice['zk_report_time'] = str(date.today()).replace('-','.')

    return advice
    