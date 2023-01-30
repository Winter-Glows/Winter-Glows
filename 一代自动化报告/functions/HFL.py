from datetime import date
from pathlib import Path

from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage, RichText

# HFL-1 CYP2C9*2    C430T   rs1799853
# HFL-2 CYP2C9*3    A1075C  rs1057910
# HFL-4 VKORC1  1639G>A rs9923231


def get_result(record: dict, CFG, mode, tpl):
    
    # record = {"genetype": {
    #   "CYP2C9-2": "CC",
    #   "CYP2C9-3": "AA",
    #   "VKORC1": "AA"
    # },}

    advice = {}

    for info in record.keys():
        exec(f"advice['{info}'] = record['{info}']")
        advice[info] = advice[info] if advice[info] else '-'

    for key in advice['genetype'].keys():
        exec(f"advice['{key}'] = advice['genetype']['{key}']")

    gene_type_HFL_1 = advice['CYP2C9-2']
    gene_type_HFL_2 = advice['CYP2C9-3']
    gene_type_HFL_4 = advice['VKORC1']

    if gene_type_HFL_1 == 'CC':
        if gene_type_HFL_2 == 'AA':
            advice['CYP2C9'] = 'CYP2C9*1/*1'
        elif gene_type_HFL_2 in ('AC','CA'):
            advice['CYP2C9'] = 'CYP2C9*1/*3'
        elif gene_type_HFL_2 == 'CC':
            advice['CYP2C9'] = 'CYP2C9*3/*3'
    if gene_type_HFL_1 in ('CT','TC'):
        if gene_type_HFL_2 == 'AA':
            advice['CYP2C9'] = 'CYP2C9*1/*2'
        elif gene_type_HFL_2 in ('AC','CA'):
            advice['CYP2C9'] = 'CYP2C9*2/*3'
    if gene_type_HFL_1 == 'TT':
        if gene_type_HFL_2 == 'AA':
            advice['CYP2C9'] = 'CYP2C9*2/*2'
    # print(advice['CYP2C9'])

    gene_type_VKORC1 = gene_type_HFL_4

    if advice['CYP2C9'][7] == advice['CYP2C9'][10]:
        gene_type_CYP2C9 = advice['CYP2C9'][:8]
        advice['note'] = f'该受检者携带两个{gene_type_CYP2C9}等位基因，VKORC1基因c.-1639位点为{gene_type_VKORC1}基因型。华法林剂量请参考下表，具体请结合临床。'
    else:
        gene_type_CYP2C9_1 = advice['CYP2C9'][:8]
        gene_type_CYP2C9_2 = advice['CYP2C9'][:6] + advice['CYP2C9'][-2:]
        advice['note'] = f'该受检者携带一个{gene_type_CYP2C9_1}等位基因与一个{gene_type_CYP2C9_2}等位基因，VKORC1基因c.-1639位点为{gene_type_VKORC1}基因型。华法林剂量请参考下表，具体请结合临床。'
    # print(advice)
    
    # 表格内容
    # CYP2C9
    advice['form11'] = 'VKORC1(c.-1639)'
    advice['form12'] = 'CYP2C9*1/*1'
    advice['form13'] = 'CYP2C9*1/*2'
    advice['form14'] = 'CYP2C9*1/*3'
    advice['form15'] = 'CYP2C9*2/*2'
    advice['form16'] = 'CYP2C9*2/*3'
    advice['form17'] = 'CYP2C9*3/*3'

    # VKORC1(c.-1639)
    advice['form21'] = 'GG'
    advice['form31'] = 'GA'
    advice['form41'] = 'AA'

    # level
    advice['form22'] = advice['form23'] = advice['form32'] = '5-7'
    advice['form24'] = advice['form25'] = advice['form26'] = advice['form33'] = advice['form34'] = advice['form35'] = advice['form42'] = advice['form43'] = '3-4'
    advice['form27'] = advice['form36'] = advice['form37'] = advice['form44'] = advice['form45'] = advice['form46'] = advice['form47'] = '0.5-2'

    if advice['CYP2C9'] == advice['form12']:
        advice['form12'] = RichText('CYP2C9*1/*1', color = 'red', bold = True)
        if gene_type_VKORC1 == advice['form21']:
            advice['form21'] = RichText('GG', color = 'red')
            advice['form22'] = RichText('5-7', color = 'red')
        elif gene_type_VKORC1 == advice['form31']:
            advice['form31'] = RichText('GA', color = 'red')
            advice['form32'] = RichText('5-7', color = 'red')
        elif gene_type_VKORC1 == advice['form41']:
            advice['form41'] = RichText('AA', color = 'red')
            advice['form42'] = RichText('3-4', color = 'red')

    if advice['CYP2C9'] == advice['form13']:
        advice['form13'] = RichText('CYP2C9*1/*2', color = 'red', bold = True)
        if gene_type_VKORC1 == advice['form21']:
            advice['form21'] = RichText('GG', color = 'red')
            advice['form23'] = RichText('5-7', color = 'red')
        elif gene_type_VKORC1 == advice['form31']:
            advice['form31'] = RichText('GA', color = 'red')
            advice['form33'] = RichText('3-4', color = 'red')
        elif gene_type_VKORC1 == advice['form41']:
            advice['form41'] = RichText('AA', color = 'red')
            advice['form43'] = RichText('3-4', color = 'red')

    if advice['CYP2C9'] == advice['form14']:
        advice['form14'] = RichText('CYP2C9*1/*3', color = 'red', bold = True)
        if gene_type_VKORC1 == advice['form21']:
            advice['form21'] = RichText('GG', color = 'red')
            advice['form24'] = RichText('3-4', color = 'red')
        elif gene_type_VKORC1 == advice['form31']:
            advice['form31'] = RichText('GA', color = 'red')
            advice['form34'] = RichText('3-4', color = 'red')
        elif gene_type_VKORC1 == advice['form41']:
            advice['form41'] = RichText('AA', color = 'red')
            advice['form44'] = RichText('0.5-2', color = 'red')

    if advice['CYP2C9'] == advice['form15']:
        advice['form15'] = RichText('CYP2C9*3/*3', color = 'red', bold = True)
        if gene_type_VKORC1 == advice['form21']:
            advice['form21'] = RichText('GG', color = 'red')
            advice['form25'] = RichText('3-4', color = 'red')
        elif gene_type_VKORC1 == advice['form31']:
            advice['form31'] = RichText('GA', color = 'red')
            advice['form35'] = RichText('3-4', color = 'red')
        elif gene_type_VKORC1 == advice['form41']:
            advice['form41'] = RichText('AA', color = 'red')
            advice['form45'] = RichText('0.5-2', color = 'red')

    if advice['CYP2C9'] == advice['form16']:
        advice['form16'] = RichText('CYP2C9*2/*3', color = 'red', bold = True)
        if gene_type_VKORC1 == advice['form21']:
            advice['form21'] = RichText('GG', color = 'red')
            advice['form26'] = RichText('3-4', color = 'red')
        elif gene_type_VKORC1 == advice['form31']:
            advice['form31'] = RichText('GA', color = 'red')
            advice['form36'] = RichText('0.5-2', color = 'red')
        elif gene_type_VKORC1 == advice['form41']:
            advice['form41'] = RichText('AA', color = 'red')
            advice['form46'] = RichText('0.5-2', color = 'red')

    if advice['CYP2C9'] == advice['form17']:
        advice['form17'] = RichText('CYP2C9*3/*3', color = 'red', bold = True)
        if gene_type_VKORC1 == advice['form21']:
            advice['form21'] = RichText('GG', color = 'red')
            advice['form27'] = RichText('0.5-2', color = 'red')
        elif gene_type_VKORC1 == advice['form31']:
            advice['form31'] = RichText('GA', color = 'red')
            advice['form37'] = RichText('0.5-2', color = 'red')
        elif gene_type_VKORC1 == advice['form41']:
            advice['form41'] = RichText('AA', color = 'red')
            advice['form47'] = RichText('0.5-2', color = 'red')

    # 图片路径
    pic_dir = Path(CFG['Dirs']['pictures'] + '/' + 'HFL')

    # 图片命名，eg: CYP2C9_2__CC.jpg
    if gene_type_HFL_1 == 'CC':
        image_path_CYP2C9_2 = pic_dir / "CYP2C9_2_CC.png"
    elif gene_type_HFL_1 in ('CT','TC'):
        image_path_CYP2C9_2 = pic_dir / "CYP2C9_2_CT.png"
    elif gene_type_HFL_1 == 'TT':
        image_path_CYP2C9_2 = pic_dir / "CYP2C9_2_TT.png"

    if gene_type_HFL_2 == 'AA':
        image_path_CYP2C9_3 = pic_dir / "CYP2C9_3_AA.png"
    elif gene_type_HFL_2 in ('AC','CA'):
        image_path_CYP2C9_3 = pic_dir / "CYP2C9_3_AC.png"
    elif gene_type_HFL_2 == 'CC':
        image_path_CYP2C9_3 = pic_dir / "CYP2C9_3_CC.png"

    if gene_type_HFL_4 == 'GG':
        image_path_VKORC1 = pic_dir / "VKORC1_GG.png"
    elif gene_type_HFL_4 in ('GA','AG'):
        image_path_VKORC1 = pic_dir / "VKORC1_GA.png"
    elif gene_type_HFL_4 == 'AA':
        image_path_VKORC1 = pic_dir / "VKORC1_AA.png"


    advice['peak_figure_CYP2C9_2'] = InlineImage(tpl, str(image_path_CYP2C9_2.absolute()), width=Mm(65), height=Mm(30))
    advice['peak_figure_CYP2C9_3'] = InlineImage(tpl, str(image_path_CYP2C9_3.absolute()), width=Mm(65), height=Mm(30))
    advice['peak_figure_VKORC1'] = InlineImage(tpl, str(image_path_VKORC1.absolute()), width=Mm(65), height=Mm(30))

    # advice['zk_entrust_time'] = record['zk_entrust_time']
    advice['zk_accept_time'] = str(advice['receive_time'][:10]).replace('-','.')
    advice['zk_report_time'] = str(date.today()).replace('-','.')
    
    return advice