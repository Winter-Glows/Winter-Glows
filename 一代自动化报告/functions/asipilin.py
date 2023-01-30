from datetime import date
from pathlib import Path
from pprint import pprint

from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage

# # normal
#     record['genetype']['PTGS1'] == 'AA'
#     record['genetype']['PEAR1'] == 'GG'
#     record['genetype']['ITGA2'] == 'CC'
#     record['genetype']['ITGB3'] == 'TT'

# # Number of variations
# 0
# 1
# 2
# 3
# 4

def get_result(record: dict, CFG, mode, tpl):  # 

    record = {'genetype':{'PTGS1':'AA',
            'PEAR1':'GG', 
            'ITGA2':'CC', 
            'ITGB3':'TT'
            },}   
     
    advice = {}

    for info in record.keys():
        exec(f"advice['{info}'] = record['{info}']")
        advice[info] = advice[info] if advice[info] else '-'

    for key in advice['genetype'].keys():
        exec(f"advice['{key}'] = advice['genetype']['{key}']")

    gene_type_PTGS1 = advice['PTGS1']
    gene_type_PEAR1 = advice['PEAR1']
    gene_type_ITGA2 = advice['ITGA2']
    gene_type_ITGB3 = advice['ITGB3']

    # 判定突变数目
    normal = {'PTGS1':'AA','PEAR1':'GG', 'ITGA2':'CC', 'ITGB3':'TT'}

    var_num = 0
    for i in normal.keys():
        if record['genetype'][i] != normal[i]:
            var_num += 1
            # print(record[i])
    # print(var_num)

    if var_num <= 1:
        advice['note'] = '该患者出现阿司匹林抵抗风险较低，可正常使用此药物，具体请结合临床。'
    elif var_num == 2:
        advice['note'] = '该患者出现阿司匹林抵抗风险较高，具体请结合临床。'
    elif var_num >= 3:
        advice['note'] = '该患者出现阿司匹林抵抗风险高，慎用或换用其它抗血小板药物。'
    # print(advice['note'])

    # 图片路径
    pic_dir = Path(CFG['Dirs']['pictures'] + '/' + 'asipilin')

    # 图片命名，eg: PTGS1_AA.jpg
    if gene_type_PTGS1 == 'AA': 
        image_path_PTGS1 = pic_dir / "PTGS1_AA.png"
    
    if gene_type_PEAR1 == 'GG':
        image_path_PEAR1 = pic_dir / "PEAR1_GG.png"
    elif gene_type_PEAR1 in ('GA','AG'):
        image_path_PEAR1 = pic_dir / "PEAR1_GA.png"
    elif gene_type_PEAR1 == 'AA':
        image_path_PEAR1 = pic_dir / "PEAR1_AA.png"

    if gene_type_ITGA2 == 'CC':
        image_path_ITGA2 = pic_dir / "ITGA2_CC.png"
    elif gene_type_ITGA2 in ('CT','TC'):
        image_path_ITGA2 = pic_dir / "ITGA2_CT.png"
    elif gene_type_ITGA2 == 'TT':
        image_path_ITGA2 = pic_dir / "ITGA2_TT.png"

    if gene_type_ITGB3 == 'TT':
        image_path_ITGB3 = pic_dir / "ITGB3_TT.png"
    elif gene_type_ITGB3 in ('TC','CT'):
        image_path_ITGB3 = pic_dir / "ITGB3_TC.png"


    advice['peak_figure_PTGS1'] = InlineImage(tpl, str(image_path_PTGS1.absolute()), width=Mm(65), height=Mm(30))
    advice['peak_figure_PEAR1'] = InlineImage(tpl, str(image_path_PEAR1.absolute()), width=Mm(65), height=Mm(30))
    advice['peak_figure_ITGA2'] = InlineImage(tpl, str(image_path_ITGA2.absolute()), width=Mm(65), height=Mm(30))
    advice['peak_figure_ITGB3'] = InlineImage(tpl, str(image_path_ITGB3.absolute()), width=Mm(65), height=Mm(30))
    # print(image_path_PTGS1)

    # advice['zk_entrust_time'] = record['zk_entrust_time']
    advice['zk_sampling_time'] = str(advice['sampling_time'][:10]).replace('-','.')
    advice['zk_accept_time'] = str(advice['receive_time'][:10]).replace('-','.')
    advice['zk_report_time'] = str(date.today()).replace('-','.')

    return advice