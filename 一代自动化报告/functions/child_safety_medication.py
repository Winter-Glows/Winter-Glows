from datetime import date
from pathlib import Path

from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage, RichText


def get_result(record: dict, CFG, mode, tpl):

    # record = {"genetype": {
    #         "12s_rRNA_1":"CT",
    #         "12s_rRNA_2":"AA",
    #         "CYP2C19-2":"GG",
    #         "CYP2C19-3":"GG",
    #         "CYP2C9-3":"CA",
    #         "CYP2D6-10":"TT",
    #         "CYP3A4-18":"TT",
    #         "CYP3A5-3":"AG"
    #         },}

    advice = {}

    for info in record.keys():
        exec(f"advice['{info}'] = record['{info}']")
        advice[info] = advice[info] if advice[info] else '-'

    for key in advice['genetype'].keys():
        exec(f"advice['{key}'] = advice['genetype']['{key}']")

    gene_type_rs4244285_2 = advice['CYP2C19_2'] = advice['CYP2C19-2']
    gene_type_rs4986893_3 = advice['CYP2C19_3'] = advice['CYP2C19-3']
    gene_type_rs1057910 = advice['CYP2C9_3'] = advice['CYP2C9-3']
    gene_type_rs1065852 = advice['CYP2D6_10'] = advice['CYP2D6-10']
    gene_type_rs28371759 = advice['CYP3A4_18'] = advice['CYP3A4-18']
    gene_type_rs776746 = advice['CYP3A5_3'] = advice['CYP3A5-3']
    gene_type_12srRNA_A1555G = advice['rRNA_2'] = advice['12s_rRNA_2']
    gene_type_12srRNA_C1494T = advice['rRNA_1'] = advice['12s_rRNA_1']

    # 药物代谢基因检测
    # CYP2C19 结果解读
    if gene_type_rs4244285_2 == 'GG':
        if gene_type_rs4986893_3 == 'GG':
            advice['result_1'] = '正常代谢'
        elif gene_type_rs4986893_3 in ('GA','AG'):
            advice['result_1'] = '中间代谢'
        elif gene_type_rs4986893_3 == 'AA':
            advice['result_1'] = '慢代谢'
    elif gene_type_rs4244285_2 in ('GA','AG'):
        if gene_type_rs4986893_3 == 'GG':
            advice['result_1'] = '中间代谢'
        elif gene_type_rs4986893_3 in ('GA','AG'):
            advice['result_1'] = '慢代谢'
    elif gene_type_rs4244285_2 == 'AA':
        if gene_type_rs4986893_3 == 'GG':
            advice['result_1'] = '慢代谢'

    # CYP2C9 结果解读
    if gene_type_rs1057910 == 'AA':
        advice['result_2'] = '正常代谢'
    elif gene_type_rs1057910 in ('AC','CA'):
        advice['result_2'] = '中间代谢'
    elif gene_type_rs1057910 == 'CC':
        advice['result_2'] = '慢代谢'

    # CYP2D6 结果解读
    if gene_type_rs1065852 == 'CC':
        advice['result_3'] = '正常代谢'
    elif gene_type_rs1065852 in ('CT','TC'):
        advice['result_3'] = '正常代谢'
    elif gene_type_rs1065852 == 'TT':
        advice['result_3'] = '慢代谢'

    # CYP3A4 结果解读
    if gene_type_rs28371759 == 'TT':
        advice['result_4'] = '正常代谢'
    elif gene_type_rs28371759 in ('TC','CT'):
        advice['result_4'] = '中间代谢'
    elif gene_type_rs28371759 == 'CC':
        advice['result_4'] = '慢代谢'

    # CYP3A5 结果解读
    if gene_type_rs776746 == 'AA':
        advice['result_5'] = '正常代谢'
    elif gene_type_rs776746 in ('AG','GA'):
        advice['result_5'] = '中间代谢'
    elif gene_type_rs776746 == 'GG':
        advice['result_5'] = '慢代谢'

    # 具体用药判定，4种结果：yes, line, wrong, down
    pic_dir = Path(CFG['Dirs']['pictures'] + '/' + 'child_safety_medication')
    image_path_yes = pic_dir / "yes.png"
    image_path_line = pic_dir / "line.png"
    image_path_wrong = pic_dir / "wrong.png"
    image_path_down = pic_dir / "down.png"

    # CYP2C19 用药指导
    if advice['result_1'] == '正常代谢':
        advice['jr_9'] = advice['jr_10'] = InlineImage(tpl, str(image_path_yes.absolute()), width=Mm(5), height=Mm(5))
        advice['xh_5'] = advice['xh_6'] = advice['xh_7'] = advice['xh_8'] = advice['xh_9'] = advice['xh_10'] = InlineImage(tpl, str(image_path_yes.absolute()), width=Mm(5), height=Mm(5))
        advice['xh_13'] = advice['xh_14'] = advice['xh_15'] = advice['xh_16'] = advice['xh_17'] = InlineImage(tpl, str(image_path_yes.absolute()), width=Mm(5), height=Mm(5))
        advice['sj_6'] = InlineImage(tpl, str(image_path_yes.absolute()), width=Mm(5), height=Mm(5))
        advice['k_21'] = InlineImage(tpl, str(image_path_yes.absolute()), width=Mm(5), height=Mm(5))
    elif advice['result_1'] == '中间代谢':
        advice['jr_9'] = advice['jr_10'] = InlineImage(tpl, str(image_path_line.absolute()), width=Mm(5), height=Mm(2))
        advice['xh_5'] = advice['xh_6'] = advice['xh_7'] = advice['xh_8'] = advice['xh_9'] = advice['xh_10'] = InlineImage(tpl, str(image_path_line.absolute()), width=Mm(5), height=Mm(2))
        advice['xh_13'] = advice['xh_14'] = advice['xh_15'] = advice['xh_16'] = advice['xh_17'] = InlineImage(tpl, str(image_path_line.absolute()), width=Mm(5), height=Mm(2))
        advice['sj_6'] = InlineImage(tpl, str(image_path_line.absolute()), width=Mm(5), height=Mm(2))
        advice['k_21'] = InlineImage(tpl, str(image_path_line.absolute()), width=Mm(5), height=Mm(2))
    elif advice['result_1'] == '慢代谢':
        advice['jr_9'] = advice['jr_10'] = InlineImage(tpl, str(image_path_wrong.absolute()), width=Mm(4), height=Mm(4))
        advice['xh_5'] = advice['xh_6'] = advice['xh_7'] = advice['xh_8'] = advice['xh_9'] = advice['xh_10'] = InlineImage(tpl, str(image_path_wrong.absolute()), width=Mm(4), height=Mm(4))
        advice['xh_13'] = advice['xh_14'] = advice['xh_15'] = advice['xh_16'] = advice['xh_17'] = InlineImage(tpl, str(image_path_wrong.absolute()), width=Mm(4), height=Mm(4))
        advice['sj_6'] = InlineImage(tpl, str(image_path_wrong.absolute()), width=Mm(4), height=Mm(4))
        advice['k_21'] = InlineImage(tpl, str(image_path_wrong.absolute()), width=Mm(4), height=Mm(4))

    # CYP2C9 用药指导
    if advice['result_2'] == '正常代谢':
        advice['jr_3'] = advice['jr_4'] = advice['jr_5'] = advice['jr_6'] = advice['jr_7'] = advice['jr_8'] = InlineImage(tpl, str(image_path_yes.absolute()), width=Mm(5), height=Mm(5))
        advice['nfm_1'] =  advice['nfm_2'] =  advice['nfm_3'] =  advice['nfm_4'] =  advice['nfm_5'] =  advice['nfm_6'] =  advice['nfm_7'] =  advice['nfm_8'] =  advice['nfm_9'] =  advice['nfm_10'] =  advice['nfm_11'] =  advice['nfm_12'] =  advice['nfm_13'] =  advice['nfm_14'] =  advice['nfm_16'] = InlineImage(tpl, str(image_path_yes.absolute()), width=Mm(5), height=Mm(5))
        advice['sj_7'] = advice['sj_8'] = advice['sj_9'] = InlineImage(tpl, str(image_path_yes.absolute()), width=Mm(5), height=Mm(5))
        advice['hx_6'] = InlineImage(tpl, str(image_path_yes.absolute()), width=Mm(5), height=Mm(5))
        advice['k_7'] = advice['k_28'] = InlineImage(tpl, str(image_path_yes.absolute()), width=Mm(5), height=Mm(5))
    elif advice['result_2'] == '中间代谢':
        advice['jr_3'] = advice['jr_4'] = advice['jr_5'] = advice['jr_6'] = advice['jr_7'] = advice['jr_8'] = InlineImage(tpl, str(image_path_line.absolute()), width=Mm(5), height=Mm(2))
        advice['nfm_1'] =  advice['nfm_2'] =  advice['nfm_3'] =  advice['nfm_4'] =  advice['nfm_5'] =  advice['nfm_6'] =  advice['nfm_7'] =  advice['nfm_8'] =  advice['nfm_9'] =  advice['nfm_10'] =  advice['nfm_11'] =  advice['nfm_12'] =  advice['nfm_13'] =  advice['nfm_14'] =  advice['nfm_16'] = InlineImage(tpl, str(image_path_line.absolute()), width=Mm(5), height=Mm(2))
        advice['sj_7'] = advice['sj_8'] = advice['sj_9'] = InlineImage(tpl, str(image_path_line.absolute()), width=Mm(5), height=Mm(2))
        advice['hx_6'] = InlineImage(tpl, str(image_path_line.absolute()), width=Mm(5), height=Mm(2))
        advice['k_7'] = advice['k_28'] = InlineImage(tpl, str(image_path_line.absolute()), width=Mm(5), height=Mm(2))
    elif advice['result_2'] == '慢代谢':
        advice['jr_3'] = advice['jr_4'] = advice['jr_5'] = advice['jr_6'] = advice['jr_7'] = advice['jr_8'] = InlineImage(tpl, str(image_path_wrong.absolute()), width=Mm(4), height=Mm(4))
        advice['nfm_1'] =  advice['nfm_2'] =  advice['nfm_3'] =  advice['nfm_4'] =  advice['nfm_5'] =  advice['nfm_6'] =  advice['nfm_7'] =  advice['nfm_8'] =  advice['nfm_9'] =  advice['nfm_10'] =  advice['nfm_11'] =  advice['nfm_12'] =  advice['nfm_13'] =  advice['nfm_14'] =  advice['nfm_16'] = InlineImage(tpl, str(image_path_wrong.absolute()), width=Mm(4), height=Mm(4))
        advice['sj_7'] = advice['sj_8'] = advice['sj_9'] = InlineImage(tpl, str(image_path_wrong.absolute()), width=Mm(4), height=Mm(4))
        advice['hx_6'] = InlineImage(tpl, str(image_path_wrong.absolute()), width=Mm(4), height=Mm(4))
        advice['k_7'] = advice['k_28'] = InlineImage(tpl, str(image_path_wrong.absolute()), width=Mm(4), height=Mm(4))

    # CYP2D6 用药指导
    if advice['result_3'] == '正常代谢':
        advice['jr_1'] = advice['jr_2'] = advice['jr_11'] = advice['jr_12'] = advice['jr_13'] = advice['jr_14'] = advice['jr_15'] = advice['jr_16'] = advice['jr_17'] = advice['jr_18'] = advice['jr_19'] = advice['jr_20'] = advice['jr_21'] = advice['jr_22'] = InlineImage(tpl, str(image_path_yes.absolute()), width=Mm(5), height=Mm(5))
        advice['xh_11'] = advice['xh_12'] = InlineImage(tpl, str(image_path_yes.absolute()), width=Mm(5), height=Mm(5))
        advice['hx_1'] = advice['hx_2'] = advice['hx_3'] = advice['hx_4'] = advice['hx_5'] = advice['hx_7'] = advice['hx_8'] = InlineImage(tpl, str(image_path_yes.absolute()), width=Mm(5), height=Mm(5))
    elif advice['result_3'] == '中间代谢':
        advice['jr_1'] = advice['jr_2'] = advice['jr_11'] = advice['jr_12'] = advice['jr_13'] = advice['jr_14'] = advice['jr_15'] = advice['jr_16'] = advice['jr_17'] = advice['jr_18'] = advice['jr_19'] = advice['jr_20'] = advice['jr_21'] = advice['jr_22'] = InlineImage(tpl, str(image_path_line.absolute()), width=Mm(5), height=Mm(2))
        advice['xh_11'] = advice['xh_12'] = InlineImage(tpl, str(image_path_line.absolute()), width=Mm(5), height=Mm(2))
        advice['hx_1'] = advice['hx_2'] = advice['hx_3'] = advice['hx_4'] = advice['hx_5'] = advice['hx_7'] = advice['hx_8'] = InlineImage(tpl, str(image_path_line.absolute()), width=Mm(5), height=Mm(2))
    elif advice['result_3'] == '慢代谢':
        advice['jr_1'] = advice['jr_2'] = advice['jr_11'] = advice['jr_12'] = advice['jr_13'] = advice['jr_14'] = advice['jr_15'] = advice['jr_16'] = advice['jr_17'] = advice['jr_18'] = advice['jr_19'] = advice['jr_20'] = advice['jr_21'] = advice['jr_22'] = InlineImage(tpl, str(image_path_wrong.absolute()), width=Mm(4), height=Mm(4))
        advice['xh_11'] = advice['xh_12'] = InlineImage(tpl, str(image_path_wrong.absolute()), width=Mm(4), height=Mm(4))
        advice['hx_1'] = advice['hx_2'] = advice['hx_3'] = advice['hx_4'] = advice['hx_5'] = advice['hx_7'] = advice['hx_8'] = InlineImage(tpl, str(image_path_wrong.absolute()), width=Mm(4), height=Mm(4))
    
    # CYP3A4 用药指导
    if advice['result_4'] == '正常代谢':
        advice['sj_3'] = advice['sj_4'] = InlineImage(tpl, str(image_path_yes.absolute()), width=Mm(5), height=Mm(5))
    elif advice['result_4'] == '中间代谢':
        advice['sj_3'] = advice['sj_4'] = InlineImage(tpl, str(image_path_line.absolute()), width=Mm(5), height=Mm(2))
    elif advice['result_4'] == '慢代谢':
        advice['sj_3'] = advice['sj_4'] = InlineImage(tpl, str(image_path_wrong.absolute()), width=Mm(4), height=Mm(4))
    
    # CYP3A5 用药指导
    if advice['result_5'] == '正常代谢':
        advice['xh_1'] = advice['xh_2'] = advice['xh_3'] = advice['xh_4'] = InlineImage(tpl, str(image_path_yes.absolute()), width=Mm(5), height=Mm(5))
        advice['sj_1'] = advice['sj_2'] = InlineImage(tpl, str(image_path_yes.absolute()), width=Mm(5), height=Mm(5))
        advice['k_8'] = advice['k_9'] = advice['k_10'] = InlineImage(tpl, str(image_path_yes.absolute()), width=Mm(5), height=Mm(5))
    elif advice['result_5'] == '中间代谢':
        advice['xh_1'] = advice['xh_2'] = advice['xh_3'] = advice['xh_4'] = InlineImage(tpl, str(image_path_line.absolute()), width=Mm(5), height=Mm(2))
        advice['sj_1'] = advice['sj_2'] = InlineImage(tpl, str(image_path_line.absolute()), width=Mm(5), height=Mm(2))
        advice['k_8'] = advice['k_9'] = advice['k_10'] = InlineImage(tpl, str(image_path_line.absolute()), width=Mm(5), height=Mm(2))
    elif advice['result_5'] == '慢代谢':
        advice['xh_1'] = advice['xh_2'] = advice['xh_3'] = advice['xh_4'] = InlineImage(tpl, str(image_path_wrong.absolute()), width=Mm(4), height=Mm(4))
        advice['sj_1'] = advice['sj_2'] = InlineImage(tpl, str(image_path_wrong.absolute()), width=Mm(4), height=Mm(4))
        advice['k_8'] = advice['k_9'] = advice['k_10'] = InlineImage(tpl, str(image_path_wrong.absolute()), width=Mm(4), height=Mm(4))

    # 抑制剂
    advice['nfm_15'] = InlineImage(tpl, str(image_path_down.absolute()), width=Mm(3), height=Mm(5))
    advice['sj_5'] = advice['sj_10'] = advice['sj_11'] = InlineImage(tpl, str(image_path_down.absolute()), width=Mm(3), height=Mm(5))
    advice['k_1'] = advice['k_2'] = advice['k_3'] = advice['k_4'] = advice['k_5'] = advice['k_6'] = InlineImage(tpl, str(image_path_down.absolute()), width=Mm(3), height=Mm(5))
    advice['k_11'] = advice['k_12'] = advice['k_13'] = advice['k_14'] = advice['k_15'] = advice['k_16'] = advice['k_17'] = advice['k_18'] = advice['k_19'] = advice['k_20'] = advice['k_21'] = InlineImage(tpl, str(image_path_down.absolute()), width=Mm(3), height=Mm(5))
    advice['k_22'] = advice['k_23'] = advice['k_24'] = advice['k_25'] = advice['k_26'] = advice['k_27'] = InlineImage(tpl, str(image_path_down.absolute()), width=Mm(3), height=Mm(5))
    advice['k_29'] = advice['k_30'] = InlineImage(tpl, str(image_path_down.absolute()), width=Mm(3), height=Mm(5))

    
    # 药物性致聋基因检测
    # 12s rRNA结果解读
    if gene_type_12srRNA_A1555G == 'AA' and gene_type_12srRNA_C1494T == 'CC':
        advice['result_6'] = '低风险'
        advice['detect'] = '未检测'
        advice['risk'] ='低风险'
        advice['describe'] = '；由于目前医学科学发展的局限性和疾病发生的复杂性，低风险结果不能完全排除个体患该病的可能性。'
    else:
        advice['result_6'] = '高风险'
        advice['detect'] = '检测'
        advice['risk'] = '高风险'
        advice['describe'] = '，使用此药物需谨慎。'
    # print(advice)

    # advice['zk_entrust_time'] = record['zk_entrust_time']
    advice['zk_accept_time'] = str(advice['receive_time'][:10]).replace('-','.')
    advice['zk_report_time'] = str(date.today()).replace('-','.')
    
    return advice