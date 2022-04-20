import docx
import win32ui
from docx import Document  # 导入读取word模块
import os
import os.path
import openpyxl
from openpyxl import *  # 导入读写excel模块
# from pythonwin import win32ui  # 导入打开文件选择框模块
import re


# 打开word文件
dlg = win32ui.CreateFileDialog(1)  # 1表示打开文件对话框
dlg.SetOFNInitialDir('C:/Users')  # 设置打开文件对话框中的初始显示目录
dlg.DoModal()
WordName = dlg.GetPathName()  # 获取选择的文件名称
WordPath = os.path.join(os.getcwd(), WordName)  # 获取选择的文件路径
doc = Document(WordPath)  # 读取word文档

# 读取word中的表格
tables = doc.tables
a = len(doc.tables)  # 获取word中的表格数量

# 打开excel文件
dlg = win32ui.CreateFileDialog(1)  # 1表示打开文件对话框
dlg.SetOFNInitialDir('C:/Users')  # 设置打开文件对话框中的初始显示目录
dlg.DoModal()
ExcelName = dlg.GetPathName()  # 获取选择的文件名称
ExcelPath = os.path.join(os.getcwd(), ExcelName)  # 获取选择的文件路径
excel = load_workbook(ExcelPath)  # 读取excel
table = excel.active  # 读取excel中的sheet1，.active为第一张表


# 定义拆分地址的函数
def cut_text(text, lenth):
    textArr = re.findall('.{' + str(lenth) + '}', text)  # 使用了re模块的findall功能
    textArr.append(text[(len(textArr) * lenth):])
    return textArr


# 判断表格数量来区分word版本
if a > 3:
    t = tables[0]
    table.cell(2, 6).value = t.cell(1, 3).text.lstrip()
    table.cell(4, 4).value = t.cell(11, 3).text.lstrip()
    table.cell(6, 4).value = t.cell(1, 3).text.lstrip()
    table.cell(10, 4).value = t.cell(5, 3).text.lstrip()
    table.cell(13, 4).value = t.cell(7, 3).text.lstrip()
    table.cell(14, 4).value = t.cell(10, 3).text.lstrip()
    table.cell(15, 4).value = t.cell(8, 3).text.lstrip()
    t2 = tables[3]
    b = len(t2.rows)
    x = [2, b - 2]
    y = [19, 19 + (b - 4)]
    # 获取和写入产品型号和数量
    for i, j in zip(x, y):
        table.cell(j, 2).value = t2.cell(i, 1).text.lstrip()
        table.cell(j, 4).value = t2.cell(i, 4).text.lstrip()
    # 字符串个数
    table.cell(6, 9).value = len(t.cell(1, 3).text.lstrip())
    table.cell(7, 9).value = len(t.cell(3, 3).text.lstrip())
    # 地址拆分
    st = cut_text(t.cell(3, 3).text, 17)
    c = len(st)
    m = [0, c - 1]
    n = [7, 7 + c - 1]
    for l, k in zip(m, n):
        table.cell(k, 4).value = st[l]


else:
    t = tables[0]
    table.cell(2, 6).value = t.cell(1, 3).text.lstrip()
    table.cell(4, 4).value = t.cell(4, 3).text.lstrip()
    table.cell(6, 4).value = t.cell(0, 1).text.lstrip()
    table.cell(10, 4).value = t.cell(4, 1).text.lstrip()
    table.cell(13, 4).value = t.cell(0, 3).text.lstrip()
    table.cell(14, 4).value = t.cell(3, 3).text.lstrip()
    table.cell(15, 4).value = t.cell(1, 3).text.lstrip()
    t2 = tables[0]
    b = len(t2.rows)
    x = [1, b - 3]
    y = [19, 19 + (b - 4)]
    for i, j in zip(x, y):
        table.cell(j, 2).value = t2.cell(i, 1).text.lstrip()
        table.cell(j, 4).value = t2.cell(i, 5).text.lstrip()
    # 字符串个数
    table.cell(6, 9).value = len(t.cell(0, 1).text.lstrip())
    table.cell(7, 9).value = len(t.cell(2, 1).text.lstrip())
    # 地址拆分
    st = cut_text(t.cell(2, 1).text, 17)
    c = len(st)
    m = [0, c - 1]
    n = [7, 7 + c - 1]
    for l, k in zip(m, n):
        table.cell(k, 4).value = st[l]
# 保存成新的文件（默认的保存位置为.py的文件位置）
excel.save(r'C:\Users\v_vwbiaoke\Desktop')
