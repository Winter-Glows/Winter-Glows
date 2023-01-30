from datetime import date
from pathlib import Path

import xlrd
from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage, RichText

# HFL-1 CYP2C9*2    C430T   rs1799853
# HFL-2 CYP2C9*3    A1075C  rs1057910
# HFL-4 VKORC1  1639G>A rs9923231


def get_result(record: dict, database, CFG, mode, tpl):
    
    # record = {"genetype": {
    #   "CYP2C9-2": "CC",
    #   "CYP2C9-3": "AA",
    #   "VKORC1": "AG"
    # },}

    advice = {}

    for info in record.keys():
        exec(f"advice['{info}'] = record['{info}']")
        advice[info] = advice[info] if advice[info] else '-'

    for key in advice['genetype'].keys():
        exec(f"advice['{key}'] = advice['genetype']['{key}']")

    # import xlrd
    # database = '../interpretations/HFL.xlsx'
    work = xlrd.open_workbook(database)
    sheet = work.sheet_by_index(0)
    rows = sheet.nrows
    row_values = {}
    for row in range(rows):
        row_values[row] = sheet.row_values(rowx = row)
        if row_values[row][0] == advice['CYP2C9-2'] and row_values[row][1] == advice['CYP2C9-3']:
            if len(row_values[row][2]) > 0:
                advice['CYP2C9'] = row_values[row][2]
                advice['note_1'] = row_values[row][3]
            else:
                advice['CYP2C9'] = '-'
                advice['note_1'] = '-'
        if row_values[row][4] == advice['VKORC1']:
            advice['note_2'] = row_values[row][5]     
    advice['note'] = advice['note_1'] + advice['note_2']
    # print(advice['CYP2C9'], advice['note'])
    
    # 表格内容
    # CYP2C9
    advice['form11'] = 'VKORC1 (c.-1639)'
    advice['form12'] = 'CYP2C9 *1/*1'
    advice['form13'] = 'CYP2C9 *1/*2'
    advice['form14'] = 'CYP2C9 *1/*3'
    advice['form15'] = 'CYP2C9 *2/*2'
    advice['form16'] = 'CYP2C9 *2/*3'
    advice['form17'] = 'CYP2C9 *3/*3'

    # VKORC1(c.-1639)
    advice['form21'] = 'GG'
    advice['form31'] = 'GA'
    advice['form41'] = 'AA'

    # level
    advice['form22'] = advice['form23'] = advice['form32'] = '5-7'
    advice['form24'] = advice['form25'] = advice['form26'] = advice['form33'] = advice['form34'] = advice['form35'] = advice['form42'] = advice['form43'] = '3-4'
    advice['form27'] = advice['form36'] = advice['form37'] = advice['form44'] = advice['form45'] = advice['form46'] = advice['form47'] = '0.5-2'

    if advice['CYP2C9'] == advice['form12']:
        advice['form12'] = RichText('CYP2C9 *1/*1', color = 'red', bold = True)
        if advice['VKORC1'] == advice['form21']:
            advice['form21'] = RichText('GG', color = 'red')
            advice['form22'] = RichText('5-7', color = 'red')
        elif advice['VKORC1'] == advice['form31']:
            advice['form31'] = RichText('GA', color = 'red')
            advice['form32'] = RichText('5-7', color = 'red')
        elif advice['VKORC1'] == advice['form41']:
            advice['form41'] = RichText('AA', color = 'red')
            advice['form42'] = RichText('3-4', color = 'red')

    if advice['CYP2C9'] == advice['form13']:
        advice['form13'] = RichText('CYP2C9 *1/*2', color = 'red', bold = True)
        if advice['VKORC1'] == advice['form21']:
            advice['form21'] = RichText('GG', color = 'red')
            advice['form23'] = RichText('5-7', color = 'red')
        elif advice['VKORC1'] == advice['form31']:
            advice['form31'] = RichText('GA', color = 'red')
            advice['form33'] = RichText('3-4', color = 'red')
        elif advice['VKORC1'] == advice['form41']:
            advice['form41'] = RichText('AA', color = 'red')
            advice['form43'] = RichText('3-4', color = 'red')

    if advice['CYP2C9'] == advice['form14']:
        advice['form14'] = RichText('CYP2C9 *1/*3', color = 'red', bold = True)
        if advice['VKORC1'] == advice['form21']:
            advice['form21'] = RichText('GG', color = 'red')
            advice['form24'] = RichText('3-4', color = 'red')
        elif advice['VKORC1'] == advice['form31']:
            advice['form31'] = RichText('GA', color = 'red')
            advice['form34'] = RichText('3-4', color = 'red')
        elif advice['VKORC1'] == advice['form41']:
            advice['form41'] = RichText('AA', color = 'red')
            advice['form44'] = RichText('0.5-2', color = 'red')

    if advice['CYP2C9'] == advice['form15']:
        advice['form15'] = RichText('CYP2C9 *3/*3', color = 'red', bold = True)
        if advice['VKORC1'] == advice['form21']:
            advice['form21'] = RichText('GG', color = 'red')
            advice['form25'] = RichText('3-4', color = 'red')
        elif advice['VKORC1'] == advice['form31']:
            advice['form31'] = RichText('GA', color = 'red')
            advice['form35'] = RichText('3-4', color = 'red')
        elif advice['VKORC1'] == advice['form41']:
            advice['form41'] = RichText('AA', color = 'red')
            advice['form45'] = RichText('0.5-2', color = 'red')

    if advice['CYP2C9'] == advice['form16']:
        advice['form16'] = RichText('CYP2C9 *2/*3', color = 'red', bold = True)
        if advice['VKORC1'] == advice['form21']:
            advice['form21'] = RichText('GG', color = 'red')
            advice['form26'] = RichText('3-4', color = 'red')
        elif advice['VKORC1'] == advice['form31']:
            advice['form31'] = RichText('GA', color = 'red')
            advice['form36'] = RichText('0.5-2', color = 'red')
        elif advice['VKORC1'] == advice['form41']:
            advice['form41'] = RichText('AA', color = 'red')
            advice['form46'] = RichText('0.5-2', color = 'red')

    if advice['CYP2C9'] == advice['form17']:
        advice['form17'] = RichText('CYP2C9 *3/*3', color = 'red', bold = True)
        if advice['VKORC1'] == advice['form21']:
            advice['form21'] = RichText('GG', color = 'red')
            advice['form27'] = RichText('0.5-2', color = 'red')
        elif advice['VKORC1'] == advice['form31']:
            advice['form31'] = RichText('GA', color = 'red')
            advice['form37'] = RichText('0.5-2', color = 'red')
        elif advice['VKORC1'] == advice['form41']:
            advice['form41'] = RichText('AA', color = 'red')
            advice['form47'] = RichText('0.5-2', color = 'red')

    # 图片路径
    pic_dir = Path(CFG['Dirs']['pictures'] + '/' + 'HFL')

    # 图片命名，eg: CYP2C9_2__CC.jpg
    gene = []
    sort = []
    for s in row_values[0][:2]:
        gene.append(s)
        sort.append(advice[s.replace('_','-')])
    gene.append(row_values[0][4])
    sort.append(advice[f'{row_values[0][4]}'])
    # print(gene, sort)

    for n in range(len(sort)):
        if sort[n]:
            exec(f"image_path_{gene[n]} = pic_dir / '{gene[n]}_{sort[n]}.png'")
            exec(f"advice['peak_figure_{gene[n]}'] = InlineImage(tpl, str(image_path_{gene[n]}.absolute()), width=Mm(65), height=Mm(30))")

    # advice['zk_entrust_time'] = record['zk_entrust_time']
    advice['zk_accept_time'] = str(advice['receive_time'][:10]).replace('-','.')
    advice['zk_report_time'] = str(date.today()).replace('-','.')
    
    return advice