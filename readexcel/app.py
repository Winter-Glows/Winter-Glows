import importlib
import json
import time
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st
from docx.opc.exceptions import PackageNotFoundError
from docxtpl import DocxTemplate

from config import CFG
from functions import jsontodict, report_win
from product import Product4Sample


def gender_output_name(project: str, CFG: dict) -> Path:
    from datetime import date
    today = date.today()
    path_str = f"{project}.output.{today.year}_{today.month}_{today.day}"
    path_str = CFG['output'] + "/" + path_str
    Path(path_str).mkdir(parents=True, exist_ok=True)
    return Path(path_str)

def get_report_info(report_name: str) -> dict:
    config = {}
    data = pd.read_csv("./project.csv", sep='\t').set_index("项目").to_dict("index")
    config['id'] = data[report_name]["编号"]
    config['program'] =  data[report_name]["程序"]
    return config

def get_project_index() -> set:
    data = pd.read_csv("./project.csv", sep='\t').set_index("项目")
    return data.index

def get_template_index() -> list:
    with open("./tpl.list",'rt', encoding='utf-8') as f:
        data = f.readlines()
    return [i.strip() for i in data]

def app():
    is_complete = False
    st.title('自动化报告')
    sidebar = st.sidebar
    project = sidebar.selectbox("项目", get_project_index())
    mode = project
    
    # uploaded_file = sidebar.file_uploader("上传文件")
    # if uploaded_file is not None:
    #     try:
    #         data = json.load(uploaded_file)
    #         st.table(data)
    #     except:
    #         raise FileNotFoundError('Only support .json !')

    info = get_report_info(project)
    result_dir = gender_output_name(project, CFG['Dirs'])
    file_name_list = []

    data_json = CFG['Dirs']['json'] + '/' + f"test_{info['program']}.json"
    database = CFG['Dirs']['interpretations'] + '/' + f"{info['program']}.xlsx"
    data = jsontodict.jsontodict(data_json)
    # res = Product4Sample().get_baseinfo('YMB20125894')  # 需指定一样本号, data['zk_id']
    # data.update(res)
    # st.write(data)

    if st.button('运行'):
        tpl = DocxTemplate(f"{CFG['Dirs']['templates']}/{mode}/{info['program']}.docx")
        # print(type(tpl))  # DocxTemplate()的类型为类，不是路径！！！
        tmp = importlib.import_module(f"functions.{info['program']}")
        context = tmp.get_result(data, database, CFG, mode, tpl)
        data.update(context)
        st.write(data)
        file_name_list.append(report_win.run(context, CFG, tpl, mode, info, result_dir))
        is_complete = True
        file_list = [result_dir / i for i in file_name_list[0]]
    
    
    if is_complete:
        st.write("生成完毕")

        with zipfile.ZipFile(f"{result_dir}.zip", mode='w', compression=zipfile.ZIP_DEFLATED) as zf:
            for file in file_list:
                zf.write(file, arcname=file.name)

        with open(f"{result_dir}.zip", "rb") as f:
            st.download_button(
            label="下载压缩包",
            data=f,
            file_name=f"{result_dir}.zip",
            mime='application/zip',
            )
    
if __name__ == "__main__":
    starttime = time.time()
    app()
    endtime = time.time()
    st.write("用时：%.6fs" %(endtime - starttime))