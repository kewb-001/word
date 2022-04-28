import openpyxl
import wc as wc
import win32ui
from docx import Document  # 导入读取word模块
import os
import os.path
import time
from openpyxl import Workbook


# 打开word文件,读取关键信息
def Handling_Information():
    dlg = win32ui.CreateFileDialog(1)  # 1代表Ture，表示打开文件对话框
    dlg.SetOFNInitialDir(r'C:\Users\14325\PycharmProjects\word\input')  # 设置打开文件对话框中的初始显示目录
    dlg.DoModal()  # 等待用户选择文件
    WordName = dlg.GetPathName()  # 获取选择的文件名称
    WordPath = os.path.join(os.getcwd(), WordName)  # 获取选择的文件路径
    print(WordPath)
    # path_word = 'input/GST-Form-SD-002 通用测试申请表.docx'
    # doc = Document(path_word)  # 读取word文档
    doc = Document(WordPath)
    tables = doc.tables  # 读取word中的表格
    table = tables[0]
    header = doc.sections[0].header.paragraphs[0].text  # 读取表头，识别哪种提取规则
    print(header)
    apply_for_company_English = table.cell(2, 4).text  # 申请公司英文名
    apply_for_company_Chinese = table.cell(3, 4).text  # 申请公司中文名
    print(apply_for_company_English)
    print(apply_for_company_Chinese)
    # path_excel = 'input/123.xlsx'
    # wb = openpyxl.load_workbook(path_excel)
    # ws = wb.active
    # sheetnames = wb.sheetnames
    # print(sheetnames)
    # sheet1 = sheetnames[0]
    # print(sheet1)
    # nrows = ws.max_row
    # ws.cell(2, 1, apply_for_company_Chinese)  # cell(行，列，值）
    # ws.cell(2, 2, apply_for_company_English)
    # wb.save(path_excel)


if __name__ == '__main__':
    Handling_Information()
