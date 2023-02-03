from datetime import date
from pathlib import Path

from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage


def get_result(record: dict, CFG, mode , tpl):  #

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
    #         "flag_hbms":Fasle,
    #         "flag_tkms":Fasle,
    #         "flag_hzxa":Fasle,}

    advice = {}
    drugs = []

    for info in record.keys():
        exec(f"advice['{info}'] = record['{info}']")
        advice[info] = advice[info] if advice[info] else '-'

    for key in advice['genetype'].keys():
        exec(f"advice['{key}'] = advice['genetype']['{key}']")

    gene_type_ABCB1_rs1045642 = advice['ABCB1_rs1045642']
    gene_type_UGT1A8 = advice['UGT1A8']
    gene_type_DROSHA = advice['DROSHA']
    gene_type_CYP1A2 = advice['CYP1A2']
    gene_type_ESR1 = advice['ESR1']
    gene_type_POR = advice['POR']
    gene_type_TPMT = advice['TPMT']
    gene_type_NUDT15 = advice['NUDT15']
    gene_type_CYP3A5_3 = advice['CYP3A5_3']
    gene_type_ABCB1_rs2032582 = advice['ABCB1_rs2032582']
    gene_type_MTHFR_1 = advice['MTHFR_1']
    gene_type_TCF7L2 = advice['TCF7L2']

    # 甲基泼尼松龙  Methylprednisolone
    if advice['flag_jjpnsl'] == True:
        drugs.append('甲基泼尼松龙')
        if gene_type_ABCB1_rs1045642 == 'TT':
            advice['tip_Methylprednisolone'] = '骨坏死的风险可能降低'
        elif gene_type_ABCB1_rs1045642 in ('TC','CT','CC'):
            advice['tip_Methylprednisolone'] = '骨坏死的风险可能增加'

    # 硫唑嘌呤  azathioprine
    if advice['flag_lcpl'] == True:
        drugs.append('硫唑嘌呤')
        if gene_type_TPMT == 'AA':
            advice['tip_azathioprine_TPMT'] = '建议按照正常剂量使用（e.g.2-3 mg/kg/day），并根据实际情况调整用药剂量'
        elif gene_type_TPMT in ('AG','GA'):
            advice['tip_azathioprine_TPMT'] = '建议减少剂量使用，并根据实际情况调整用药剂量'
        elif gene_type_TPMT == 'GG':
            advice['tip_azathioprine_TPMT'] = '建议减少剂量使用或换用其它药物'
    
        if gene_type_NUDT15 == 'CC':
            advice['tip_azathioprine_NUDT15'] = '发生白细胞减少、中性粒细胞减少或脱发的风险降低，可按照正常剂量使用，并根据实际情况调整用药剂量'
        elif gene_type_NUDT15 in ('CT','TC'):
            advice['tip_azathioprine_NUDT15'] = '发生白细胞减少、中性粒细胞减少或脱发的风险增加，可减少剂量使用，并根据实际情况调整用药剂量'
        elif gene_type_NUDT15 == 'TT':
            advice['tip_azathioprine_NUDT15'] = '发生白细胞减少、中性粒细胞减少或脱发的风险增加，可减少剂量或换用其它药物使用'
    
    # 霉酚酸酯  mycophenolatemofetil
    if advice['flag_mfsz'] == True:
        drugs.append('霉酚酸酯')
        if gene_type_UGT1A8 == 'CC':
            advice['tip_mycophenolatemofetil_UGT1A8'] = '出现腹泻的风险可能增加'
        elif gene_type_UGT1A8 in ('CT','TC','TT'):
            advice['tip_mycophenolatemofetil_UGT1A8'] = '出现腹泻的风险可能降低'

        if gene_type_ABCB1_rs2032582 == 'GG':
            advice['tip_mycophenolatemofetil_ABCB1'] = '出现急性排斥反应风险可能降低'
        elif gene_type_ABCB1_rs2032582 in ('AG','GT','AA','AT','TT'):
            advice['tip_mycophenolatemofetil_ABCB1'] = '出现急性排斥反应风险可能增加'

    # 强的松    prednisone
    if advice['flag_qds'] == True:
        drugs.append('强的松')
        if gene_type_DROSHA in ('GG','GA','AG'):
            advice['tip_prednisone_DROSHA'] = '药物毒副作用可能降低'
        elif gene_type_DROSHA == 'AA':
            advice['tip_prednisone_DROSHA'] = '药物毒副作用可能增加'
        
        if gene_type_ABCB1_rs1045642 in ('TT','TC','CT'):
            advice['tip_prednisone_ABCB1'] = '药物应答可能较好'
        elif gene_type_ABCB1_rs1045642 == 'CC':
            advice['tip_prednisone_ABCB1'] = '药物应答可能较差'
    
    # 来氟米特  Leflunomide
    if advice['flag_lfmt'] == True:
        drugs.append('来氟米特')
        if gene_type_CYP1A2 in ('AA','AC','CA'):
            advice['tip_Leflunomide_CYP1A2'] = '药物毒副作用可能降低'
        elif gene_type_CYP1A2 == 'CC':
            advice['tip_Leflunomide_CYP1A2'] = '药物毒副作用可能增加'

        if gene_type_ESR1 in ('CC','TC','CT'):
            advice['tip_Leflunomide_ESR1'] = '药物应答可能较差'
        elif gene_type_ESR1 == 'TT':
            advice['tip_Leflunomide_ESR1'] = '药物应答可能较好'

    # 环孢霉素  cyclosporine
    if advice['flag_hbms'] == True:
        drugs.append('环孢霉素')
        if gene_type_CYP3A5_3 in ('AA','AG','GA'):
            advice['tip_cyclosporine_CYP3A5_3'] = '可能需要较高剂量'
        elif gene_type_CYP3A5_3 == 'GG':
            advice['tip_cyclosporine_CYP3A5_3'] = '可能需要较低剂量'

        if gene_type_TCF7L2 in ('CC','CT','TC'):
            advice['tip_cyclosporine_TCF7L2'] = '新发糖尿病的可能性降低'
        elif gene_type_TCF7L2 == 'TT':
            advice['tip_cyclosporine_TCF7L2'] = '新发糖尿病的可能性增加'
        
        if gene_type_ABCB1_rs2032582 in ('GG','AG', 'GT'):
            advice['tip_cyclosporine_ABCB1'] = '药物耐药风险可能降低'
        elif gene_type_ABCB1_rs2032582 in ('AA', 'AT', 'TT'):
            advice['tip_cyclosporine_ABCB1'] = '药物耐药风险可能增加'

    # 他克莫司  Tacrolimus
    if advice['flag_tkms'] == True:
        drugs.append('他克莫司')
        if gene_type_CYP3A5_3 in ('AA','AG','GA'):
            advice['tip_Tacrolimus_CYP3A5_3'] = '建议增加药物起始剂量（为标准剂量的1.5-2倍，但不超过0.3mg/公斤体重/天），同时需结合其他临床因素调整剂量'
        elif gene_type_CYP3A5_3 == 'GG':
            advice['tip_Tacrolimus_CYP3A5_3'] = '建议按照药物起始剂量服用，并请结合其他临床因素调整剂量'
        
        if gene_type_POR == 'CC':
            advice['tip_Tacrolimus_POR'] = '患新发糖尿病的风险可能降低'
        elif gene_type_POR in ('TT','CT','TC'):
            advice['tip_Tacrolimus_POR'] = '患新发糖尿病的风险可能增加'

    # 环磷酰胺  cyclophosphamide
    if advice['flag_hlxa'] == True:
        drugs.append('环磷酰胺')
        if gene_type_MTHFR_1 in ('CC','CT','TC'):
            advice['tip_cyclophosphamide_MTHFR_1'] = '药物毒副作用可能降低'
        elif gene_type_MTHFR_1 == 'TT':
            advice['tip_cyclophosphamide_MTHFR_1'] = '药物毒副作用可能增加'
    # print(advice)

    drugs = '、'.join(drugs)
    advice['drugs'] = drugs

    # advice['zk_entrust_time'] = record['zk_entrust_time']
    advice['zk_accept_time'] = str(advice['receive_time'][:10]).replace('-','.')
    advice['zk_report_time'] = str(date.today()).replace('-','.')

    return advice