# -*- coding: utf-8 -*-
"""
Spyder编辑器

这是一个临时脚本文件
"""


import xlrd,xlwt
import os
import random

def read_excel():
    
    filename ='C:/Users/15193/Desktop/GRE.xls'
    filePath = os.path.join(os.getcwd(), filename)
    
    data = xlrd.open_workbook(filePath)
    table = data.sheet_by_index(0)

    nrows =table.nrows
    Sword= []
    Smean= []

    for i in range(0,nrows-1):
        Sword.append(table.cell(i,0).value)
        Smean.append(table.cell(i,1).value)
    return Sword,Smean

def write_excel(file_max,file_num,Sword,Smean):

    wb = xlwt.Workbook()
    for i in range(0,file_num-1):
        sheets_name=['day'+str(i+1) for i in range(file_num)]
        sheet=sheets_name[i]
        sheet= wb.add_sheet(sheets_name[i], cell_overwrite_ok=True)
        sheet.col(0).width=2500
        sheet.col(1).width=6000
        sheet.col(2).width=40000
        sheet.write(0,0,'忘记次数')
        sheet.write(0,1,'单词')
        sheet.write(0,2,'意思解释')
        sheet.write(0,3,'随机排序')
        for j in range(0,file_max-1):
            sheet.write(j+1,0,0)
            sheet.write(j+1,1, Sword[file_max*i+j]) # 写入Sword
            sheet.write(j+1,2, Smean[file_max*i+j]) # 写入Smean
            sheet.write(j+1,3,random.random())
    wb.save(r'C:/Users/15193/Desktop/GRE_add.xls') # 保存文件
    
def main():
    Sword,Smean=read_excel()
    write_excel(100,31,Sword,Smean)
    
if __name__=="__main__":
    main()