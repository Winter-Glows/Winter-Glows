import PySimpleGUI as sg

from pdf2docx import Converter


def gui():

    layout = [  [sg.FileBrowse('请选择PDF文件:', font = ("楷体", 14), key = 'pdf'), 
                 sg.Text('', key = 'pdf', size = (31,5))],
                [sg.Text('请输入起始页码：'), sg.Input('', key ='start', size = (7,1)),
                 sg.Text('请输入结束页码：'), sg.Input('', key ='end', size = (7,1))], 
                [sg.Text('请输入需要指定转换的页码（eg:4.6）：'), sg.Input(key ='list', size = (14,1))],
                [sg.Text('起始页码默认为1，结束页码默认末页页码！\n默认全页码转换！', text_color = 'pink')],
                [sg.Output(size=(50, 15))],
                [sg.OK('运 行', button_color = 'blue'), sg.Quit('退 出', button_color = 'red')]
    ]

    window = sg.Window('PDF_DOCX转换器', layout, font=("楷体", 12))
    
    while True:
        event, values = window.read()

        if '.' in values['list']:
            split = values['list'].split('.')
            list = []
            for i in split:
                list.append(int(i))
        else:
            list = values['list']

        if event in (sg.WIN_CLOSED, '退 出'):
            break
        elif event == '运 行':
            if values['pdf']:
                if values['pdf'].endswith('.pdf'):
                    print('转换起始页码：', values['start'])
                    print('转换终止页码：', values['end'])
                    print('需要特定转换的页码：', list)
                    pdftodocx(values['pdf'], values['start'], values['end'], list)
                else:
                    print('仅支持.pdf格式文件！\n请重新选择：')
            else:
                print('请先选择文件！')
        
    window.close()

def pdftodocx(pdf_path: str, start_page: str, end_page: str, list_pages):
    cv = Converter(pdf_path)

    if list_pages:
        if start_page or end_page:
            print('起始终止页码不能与特定转换页码同时存在！')
        else:
            list_pages_py = list_pages[:]
            if type(list_pages) is list:
                for i in range(len(list_pages)):
                    list_pages[i] = list_pages[i] - 1
            elif type(list_pages) is str:
                list_pages = str(int(list_pages)-1)
            docx_path = pdf_path.replace('.pdf',f"_N_N_{list_pages_py}.docx")
            print('{0}正在将pdf文件转换为docx文件{0}'.format('*'*10))
            cv.convert(docx_path, pages = list_pages)
            print('\ndocx文件所在路径为：\n', docx_path)
            print('\n{0}docx文件转换完毕{0}'.format('*'*10))
    else:
        print('{0}正在将pdf文件转换为docx文件{0}'.format('*'*10))
        if start_page:
            if end_page:
                docx_path = pdf_path.replace('.pdf',f"_{start_page}_{end_page}_N.docx")
                cv.convert(docx_path, start = str(int(start_page)-1), end = end_page)
            else:
                docx_path = pdf_path.replace('.pdf',f"_{start_page}_L_N.docx")
                cv.convert(docx_path, start = str(int(start_page)-1))
        elif end_page:
            if start_page:
                docx_path = pdf_path.replace('.pdf',f"_{start_page}_{end_page}_N.docx")
                cv.convert(docx_path, start = str(int(start_page)-1), end = end_page)
            else:
                docx_path = pdf_path.replace('.pdf',f"_S_{end_page}_N.docx")
                cv.convert(docx_path, end = end_page)
        else:
            docx_path = pdf_path.replace('.pdf',f"_L_L_N.docx")
            cv.convert(docx_path)
        print('\ndocx文件所在路径为：\n', docx_path)
        print('\n{0}docx文件转换完毕{0}'.format('*'*10))

    cv.close()

if __name__ == '__main__':
    gui()         


# from pdf2docx import Converter

# pdf_file = '1.pdf'
# docx_file = '1.docx'

# # convert pdf to docx
# cv = Converter(pdf_file)
# cv.convert(docx_file) # 默认参数start=0, end=None
# cv.close()

# more samples
# cv.convert(docx_file, start=1) # 转换第2页到最后一页
# cv.convert(docx_file, pages=[1,3,5]) # 转换第2，4，6页