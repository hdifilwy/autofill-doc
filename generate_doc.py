import pandas as pd
from docxtpl import DocxTemplate
import shutil
import os

print("请把Word模版文件和Excel数据文件放在同一个目录下！")
tpl = input("请输入模版文件名（包含.docx扩展名。直接回车使用默认模版 template.docx）")
if not tpl:
    tpl = "template.docx"
if not os.path.isfile(tpl.strip()):
    print("Word模版不存在！")
    exit()

excel = input("请输入Excel 文件名（包含.xlsx扩展名）：")
if not os.path.isfile(excel.strip()):
    print("Excel 数据文件不存在")
    exit()

org = input("请输入委托单位：")

df = pd.read_excel(excel.strip(), index_col="基站编号")
# print(df.head(3))

output_folder = "output"
if os.path.exists(output_folder):
    print(f"删除'{output_folder}'")
    shutil.rmtree(output_folder)
os.makedirs(output_folder)

for idx, row in df.iterrows():
    print("基站编号：{0},  基站名称：{1}".format(idx, row["基站名称"]))
    doc = DocxTemplate(tpl.strip())
    doc.render(dict(
        基站名称 = row["基站名称"],
        委托单位 = org.strip(),
        基站地址 = row["基站站址"],
        经度    = row["经度"],
        纬度    = row["纬度"],
        天线高度    = row["天线高度（米）"]
    ))
    name= row["基站名称"]
    doc.save(f"./output/{name}.docx")

print(f"执行完毕，请参见{output_folder}文件夹内容。")