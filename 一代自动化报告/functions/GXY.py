from datetime import date
from pathlib import Path

from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage


def get_result(record: dict, CFG, mode, tpl):

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

    gene_type_CYP2C9 = advice['CYP2C9']
    gene_type_AGTR1 = advice['AGTR1']
    gene_type_ACE = advice['ACE']
    gene_type_CYP2D6 = advice['CYP2D6']
    gene_type_ADRB1 = advice['ADRB1']
    gene_type_CYP3A5 = advice['CYP3A5']
    gene_type_ADD1 = advice['ADD1']
    gene_type_NEDD4L = advice['NEDD4L']

    # 1. 血管紧张素Ⅱ受体拮抗剂 CYP2C9(c.1075A>C) AGTR1(c.*86A>C)
    if gene_type_CYP2C9 == 'AA':
        if gene_type_AGTR1 == 'AA':
            advice['text_1'] = '血管紧张素Ⅱ受体拮抗剂：CYP2C9基因型为AA，代谢能力正常，建议使用常规剂量；AGTR1基因型为AA，与AC或CC相比，体液和肾血流动力学反应较差，建议适当增加剂量或换药。'
            advice['tip_1'] = '可用(↑)'
        elif gene_type_AGTR1 in ('AC','CA','CC'):
            advice['text_1'] = f'血管紧张素Ⅱ受体拮抗剂：CYP2C9基因型为AA，代谢能力正常，建议使用常规剂量；AGTR1基因型为{gene_type_AGTR1}，与AA相比，体液和肾血流动力学反应正常，建议使用常规剂量。'
            advice['tip_1'] = '推荐'
    if gene_type_CYP2C9 in ('AC','CA'):
        if gene_type_AGTR1 == 'AA':
            advice['text_1'] = f'血管紧张素Ⅱ受体拮抗剂：CYP2C9基因型为{gene_type_CYP2C9}，代谢能力降低，建议适当降低剂量或换药。AGTR1基因型为AA，与AC或CC相比，体液和肾血流动力学反应较差，建议适当增加剂量或换药。'
            advice['tip_1'] = '慎用'
        elif gene_type_AGTR1 in ('AC','CA','CC'):
            advice['text_1'] = f'血管紧张素Ⅱ受体拮抗剂：CYP2C9基因型为{gene_type_CYP2C9}，代谢能力降低，建议适当降低剂量或换药。AGTR1基因型为{gene_type_AGTR1}，与AA相比，体液和肾血流动力学反应正常，建议使用常规剂量。'
            advice['tip_1'] = '可用(↓)'
    if gene_type_CYP2C9 == 'CC':
        if gene_type_AGTR1 == 'AA':
            advice['text_1'] = '血管紧张素Ⅱ受体拮抗剂：CYP2C9基因型为CC，代谢能力降低，建议适当降低剂量或换药。AGTR1基因型为AA，与AC或CC相比，体液和肾血流动力学反应较差，建议适当增加剂量或换药。'
            advice['tip_1'] = '慎用'
        elif gene_type_AGTR1 in ('AC','CA','CC'):
            advice['text_1'] = f'血管紧张素Ⅱ受体拮抗剂：CYP2C9基因型为CC，代谢能力降低，建议适当降低剂量或换药。AGTR1基因型为{gene_type_AGTR1}，与AA相比，体液和肾血流动力学反应正常，建议使用常规剂量。'
            advice['tip_1'] = '可用(↓)'
        
    # 2. 血管紧张素转换酶抑制剂 ACE
    if gene_type_ACE == 'II':
        advice['text_2'] = '血管紧张素转换酶抑制剂：ACE基因型为II，药物敏感性差，建议增加剂量或换药。'
        advice['tip_2'] = '可用(↑)'
    elif gene_type_ACE == 'ID':
        advice['text_2'] = '血管紧张素转换酶抑制剂：ACE基因型为ID，药物敏感性差，建议增加剂量或换药。'
        advice['tip_2'] = '可用(↑)'
    elif gene_type_ACE == 'DD':
        advice['text_2'] = '血管紧张素转换酶抑制剂：ACE基因型为DD，药物敏感性好，建议使用该类药物。'
        advice['tip_2'] = '推荐'

    # 3. β受体阻断药 CYP2D6(c.100C>T) ADRB1(c.1165G>C)
    if gene_type_CYP2D6 == 'CC':
        if gene_type_ADRB1 in ('GG','GC','CG'):
            advice['text_3'] = f'β受体阻断药：CYP2D6基因型为CC，药物代谢能力较高，建议使用该类药物。ADRB1基因型为{gene_type_ADRB1}，药物敏感性较差，建议适当增加剂量或换药。'
            advice['tip_3'] = '可用(↑)'
        elif gene_type_ADRB1 == 'CC':
            advice['text_3'] = 'β受体阻断药：CYP2D6基因型为CC，药物代谢能力较高，建议使用该类药物。ADRB1基因型为CC，药物敏感性较好，建议使用该类药物。'
            advice['tip_3'] = '推荐'
    if gene_type_CYP2D6 in ('CT','TC','TT'):
        if gene_type_ADRB1 in ('GG','GC','CG'):
            advice['text_3'] = f'β受体阻断药：CYP2D6基因型为{gene_type_CYP2D6}，药物代谢能力降低，建议适当降低剂量或换药。ADRB1基因型为{gene_type_ADRB1}，药物敏感性较差，建议适当增加剂量或换药。'
            advice['tip_3'] = '慎用'
        elif gene_type_ADRB1 == 'CC':
            advice['text_3'] = f'β受体阻断药：CYP2D6基因型为{gene_type_CYP2D6}，药物代谢能力降低，建议适当降低剂量或换药。ADRB1基因型为CC，药物敏感性较好，建议使用该类药物。'
            advice['tip_3'] = '可用(↓)'

    # 4. 钙拮抗剂 CYP3A5(c.219-237A>G)
    if gene_type_CYP3A5 == 'AA':
        advice['text_4'] = '钙拮抗剂：CYP3A5基因型为AA，药物代谢能力正常，建议使用该类药物。'
        advice['tip_4'] = '推荐'
    elif gene_type_CYP3A5 in ('GA','AG','GG'):
        advice['text_4'] = f'钙拮抗剂：CYP3A5基因型为{gene_type_CYP3A5}，药物代谢能力较高，药效降低，建议适当增加剂量或换药。'
        advice['tip_4'] = '可用(↑)'

    # 5. 利尿药 ADD1(c.1378G>T) NEDD4L(-326G>A)
    if gene_type_ADD1 == 'GG':
        if gene_type_NEDD4L == 'GG':
            advice['text_5'] = '利尿剂：ADD1基因型为GG，药物敏感性较差，建议适当增加剂量或换药。NEDD4L基因型为GG，药物敏感性较差，建议适当增加剂量或换药。'
            advice['tip_5'] = '慎用'
        elif gene_type_NEDD4L in ('GA','AG','AA'):
            advice['text_5'] = f'利尿剂：ADD1基因型为GG，药物敏感性较差，建议适当增加剂量或换药。NEDD4L基因型为{gene_type_NEDD4L}，药物敏感性较好，建议使用该类药物。'
            advice['tip_5'] = '慎用'
    if gene_type_ADD1 in ('GT','TG','TT'):
        if gene_type_NEDD4L == 'GG':
            advice['text_5'] = f'利尿剂：ADD1基因型为{gene_type_ADD1}，药物敏感性较好，建议使用该类药物。NEDD4L基因型为GG，药物敏感性较差，建议适当增加剂量或换药。'
            advice['tip_5'] = '可用(↑)'
        elif gene_type_NEDD4L in ('GA','AG','AA'):
            advice['text_5'] = f'利尿剂：ADD1基因型为{gene_type_ADD1}，药物敏感性较好，建议使用该类药物。NEDD4L基因型为{gene_type_NEDD4L}，药物敏感性较好，建议使用该类药物。'
            advice['tip_5'] = '推荐'
    # print(advice)

    # advice['zk_entrust_time'] = record['zk_entrust_time']
    advice['zk_accept_time'] = str(advice['receive_time'][:10]).replace('-','.')
    advice['zk_report_time'] = str(date.today()).replace('-','.')
    
    return advice