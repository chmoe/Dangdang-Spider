# -*- coding: UTF-8 -*-

import pandas as pd
import numpy as np
import xlwt  # 用于写入Excel
from xlutils.copy import copy
from xlrd import open_workbook  # 打开Excel进行二次写入
import os
import win32com.client as win32


filename = "StockInfo.xls" # 储存文件名
workbook = xlwt.Workbook()
initSheetName = "sheet1"
sheet = workbook.add_sheet(initSheetName) # 创建一个空的sheet，否则不能保存文件
workbook.save(filename)                           # 这里是文件名字
book_stock_info = []
def get():
    count = 300

    data_frame=pd.read_excel('./Excel_test1.xls') # 读取书籍信息

    book_id = data_frame['id']
    book_sell_price = data_frame['售价']
    # print(len(book_sell_price))
    select_id = np.unique(np.random.randint(0, len(book_id), count)) # 随机选取count个不重复的书籍信息

    books_count = np.random.randint(0,180, len(book_id))

    book_zip_info = list(zip(book_id, books_count, book_sell_price))
    

    for i in select_id:
        book_stock_info.append(book_zip_info[i])
    print(book_stock_info)



def kore():
    

    rexcel = open_workbook(filename) # 用wlrd提供的方法读取一个excel文件
    excel = copy(rexcel) # 用xlutils提供的copy方法将xlrd的对象转化为xlwt的对象
    table = excel.get_sheet(initSheetName) # 用xlwt对象的方法获得要操作的sheet
    # excel.save(filename) # xlwt对象的保存方法，这时便覆盖掉了原来的excel
    sheet = table
    sheet.write(0,0,"book_id")
    sheet.write(0,1,"库存量")
    sheet.write(0,2,"售价")

    count = 1
    for item in book_stock_info:
        sheet.write(count,0, str(item[0])) # bookid
        sheet.write(count,1, int(item[1])) # stock
        sheet.write(count,2, str(item[2])) # price
        count += 1
    excel.save(filename) # xlwt对象的保存方法，这时便覆盖掉了原来的excel

def turn2XLSX(filename):
    fname = os.getcwd()+"\\"+filename
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(fname)

    wb.SaveAs(fname+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
    wb.Close()                               #FileFormat = 56 is for .xls extension
    excel.Application.Quit()
def start():
    kore()
    os.system("cls")
    print("正在处理文件中......")
    turnFileName = filename+'x'
    if os.path.exists(turnFileName): # 删除现有文件
        os.remove(turnFileName)
    turn2XLSX(filename) # 转换文件格式

def main():
    get()
    start()
    os.system('cls')
    print("已完成")

if __name__=="__main__":
    main()
