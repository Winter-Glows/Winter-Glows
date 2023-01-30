from datetime import date
from pathlib import Path
from pprint import pprint

import xlrd
from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage


def get_result(record: dict, database, CFG, mode, tpl):  # 

    # record = {'genetype':{'PTGS1':'AA',
    #         'PEAR1':'AA', 
    #         'ITGA2':'CT', 
    #         'ITGB3':'TT'
    #         },}   

    advice = {}
    for info in record.keys():
        exec(f"advice['{info}'] = record['{info}']")
        advice[info] = advice[info] if advice[info] else '-'
    for key in advice['genetype'].keys():
        exec(f"advice['{key}'] = advice['genetype']['{key}']")

    # import xlrd
    # database = '../interpretations/asipilin.xlsx'
    work = xlrd.open_workbook(database)
    sheet = work.sheet_by_index(0)
    rows = sheet.nrows
    row_values = {}
    gene_type = {}
    for row in range(rows):
        row_values[row] = sheet.row_values(rowx = row)
        gene_type[row] = row_values[row][:4]
    gene_type.pop(0)
    # print(gene_type)

    gene = []
    sort = []
    for s in row_values[0][:4]:
        gene.append(s)
        sort.append(advice[s])
    # print(gene, sort)

    for i, j in gene_type.items():
        if j == sort:
            advice['note'] = row_values[i][4]
    # print(advice['note'])

    # 图片路径
    pic_dir = Path(CFG['Dirs']['pictures'] + '/' + 'asipilin')

    # 图片命名，eg: PTGS1_AA.jpg
    for n in range(len(sort)):
        if sort[n]:
            exec(f"image_path_{gene[n]} = pic_dir / '{gene[n]}_{sort[n]}.png'")
            exec(f"advice['peak_figure_{gene[n]}'] = InlineImage(tpl, str(image_path_{gene[n]}.absolute()), width=Mm(65), height=Mm(30))")
    # print(image_path_PTGS1)

    # advice['zk_entrust_time'] = record['zk_entrust_time']
    advice['zk_sampling_time'] = str(advice['sampling_time'][:10]).replace('-','.')
    advice['zk_accept_time'] = str(advice['receive_time'][:10]).replace('-','.')
    advice['zk_report_time'] = str(date.today()).replace('-','.')

    return advice