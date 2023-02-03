from datetime import date
from pathlib import Path
from pprint import pprint

from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage


def get_result(record: dict, CFG, mode, tpl):
    # record = {'genetype': {"TNB-10n": "TT",
    #           "TTL-1_1": "TT",
    #           "TTL-1_2": "TT"
    #         },}

    advice = {}

    for info in record.keys():
        exec(f"advice['{info}'] = record['{info}']")
        # pprint(advice)
        advice[info] = advice[info] if advice[info] else '-'

    for key in advice['genetype'].keys():
        exec(f"advice['{key}'] = advice['genetype']['{key}']")

    gene_type_SLCO1B1 = advice['SLCO1B1'] = advice['TNB-10n']
    gene_type_apoe_T388C = advice['apoe_T388C'] = advice['TTL-1_1']
    gene_type_apoe_C526T = advice['apoe_C526T'] = advice['TTL-1_2']

    pic_dir = Path(CFG['Dirs']['pictures'] + '/' + 'statins')

    # SLCO1B1【他汀与肌病风险】 图片
    if gene_type_SLCO1B1 == 'TT':
        advice['result_SLCO1B1'] = '您的检测结果为SLCO1B1*1/*1（TT），可以按正常剂量使用他汀类药物。'
        iamge_path_SLCO1B1 = pic_dir / "SLCO1B1_TT.png"
    elif gene_type_SLCO1B1 in ('CT','TC'):
        advice['result_SLCO1B1'] = f'您的检测结果为SLCO1B1*1/*1（{gene_type_SLCO1B1}），建议降低用药剂量或换用其它降脂药物,并加强临床关注。'
        iamge_path_SLCO1B1 = pic_dir / "SLCO1B1_TC.png"
    elif gene_type_SLCO1B1 == 'CC':
        advice['result_SLCO1B1'] = '您的检测结果为SLCO1B1*1/*1（CC），对于携带SLCO1B1纯合突变需将用药剂量调整为常规剂量的1/4或者更低。'
        iamge_path_SLCO1B1 = pic_dir / "SLCO1B1_CC.png"
    # print(advice['SLCO1B1'], advice['result_SLCO1B1'])

    # ApoE【他汀与降脂效果】    图片
    if gene_type_apoe_T388C == 'TT':
        iamge_path_apoe_T388C = pic_dir / "apoe_T388C_TT.png"
        if gene_type_apoe_C526T == 'TT':
            advice['result_apoe'] = '您的检测结果为ApoE3（ε2/ε2），他汀类药物治疗疗效较好。'
        elif gene_type_apoe_C526T in ('TC','CT'):
            advice['result_apoe'] = '您的检测结果为ApoE3（ε2/ε3），他汀类药物治疗疗效较好。'
        elif gene_type_apoe_C526T == 'CC':
            advice['result_apoe'] = '您的检测结果为ApoE3（ε3/ε3），他汀类药物治疗疗效正常。'
    if gene_type_apoe_T388C in ('TC','CT'):
        iamge_path_apoe_T388C = pic_dir / "apoe_T388C_TC.png"
        if gene_type_apoe_C526T in ('CT','TC'):
            advice['result_apoe'] = '您的检测结果为ApoE3（ε2/ε4），他汀类药物治疗疗效正常。'
        elif gene_type_apoe_C526T == 'CC':
            advice['result_apoe'] = '您的检测结果为ApoE3（ε3/ε4），他汀类药物治疗疗效较差。'
    if gene_type_apoe_T388C == 'CC':
        iamge_path_apoe_T388C = pic_dir / "apoe_T388C_CC.png"
        if gene_type_apoe_C526T == 'CC':
            advice['result_apoe'] = '您的检测结果为ApoE3（ε4/ε4），他汀类药物治疗疗效较差。'
    # print(advice['result_apoe'])


    # 图片命名，eg: CYP2C9_2__CC.jpg
    if gene_type_apoe_C526T == 'TT':
        iamge_path_apoe_C526T = pic_dir / "apoe_C526T_TT.png"
    elif gene_type_apoe_C526T in ('CT','TC'):
        iamge_path_apoe_C526T = pic_dir / "apoe_C526T_CT.png"
    elif gene_type_apoe_C526T == 'CC':
        iamge_path_apoe_C526T = pic_dir / "apoe_C526T_CC.png"

    advice['peak_figure_SLCO1B1'] = InlineImage(tpl, str(iamge_path_SLCO1B1.absolute()), width=Mm(45), height=Mm(27))
    advice['peak_figure_apoe_T388C'] = InlineImage(tpl, str(iamge_path_apoe_T388C.absolute()), width=Mm(45), height=Mm(27))
    advice['peak_figure_apoe_C526T'] = InlineImage(tpl, str(iamge_path_apoe_C526T.absolute()), width=Mm(45), height=Mm(27))

    # advice['zk_entrust_time'] = record['zk_entrust_time']
    advice['zk_accept_time'] = str(advice['receive_time'][:10]).replace('-','.')
    advice['zk_report_time'] = str(date.today()).replace('-','.')
    
    return advice