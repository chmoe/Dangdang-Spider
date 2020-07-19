# -*- coding: utf-8 -*-

import os
import urllib.request  # 用去获取网站链接请求
from urllib.parse import quote  # 用于中文的URL编码转换

import openpyxl
import win32com.client as win32
import xlwt  # 用于写入Excel
from bs4 import BeautifulSoup  # 用于读取网站内容
from xlrd import open_workbook  # 打开Excel进行二次写入
from xlutils.copy import copy
from fake_useragent import UserAgent
import time


filename = "Excel_test1.xls" # 储存文件名
workbook = xlwt.Workbook()
sheet = workbook.add_sheet(" ") # 创建一个空的sheet，否则不能保存文件
workbook.save(filename)                            # 这里是文件名字

def findISBN(s):
    findContent = "国际标准书号ISBN："
    isbnStartPosition = s.rfind(findContent)+len(findContent)
    isbn = s[isbnStartPosition:isbnStartPosition+13]
    return isbn
def findId(s):
    findContent = "http://product.dangdang.com/"
    idStartPosition = s.rfind(findContent)+len(findContent)
    idNumber = s[idStartPosition:-5]
    return idNumber
def kore(keyword,ranges = 1, flag = False):
    '''
    按照关键词查找书籍信息的核心代码
    '''
    rexcel = open_workbook(filename) # 用wlrd提供的方法读取一个excel文件
    excel = copy(rexcel) # 用xlutils提供的copy方法将xlrd的对象转化为xlwt的对象
    if(flag):table = excel.get_sheet(" ") # 用xlwt对象的方法获得要操作的sheet
    else:table = excel.add_sheet(keyword) # 用xlwt对象的方法获得要操作的sheet
    excel.save(filename) # xlwt对象的保存方法，这时便覆盖掉了原来的excel
    sheet = table
    sheet.write(0,0,"id")
    sheet.write(0,1,"题目")
    sheet.write(0,2,"ISBN")
    sheet.write(0,3,"作者")
    sheet.write(0,4,"定价")
    sheet.write(0,5,"售价")
    sheet.write(0,6,"出版社")
    sheet.write(0,7,"类别")
    
    if(flag):count = rexcel.sheets()[0].nrows # 用wlrd提供的方法获得现在已有的行数
    else:count = 1
    # ranges = 1 # 查找的页面最大值
    icount = 1
    for i in range(1, ranges+1):
        url = "http://search.dangdang.com/?key={}&act=input&page_index={}".format(quote(keyword, 'utf-8'), i)

        f = urllib.request.urlopen(url)
        html = f.read().decode('gb18030')
        # .decode('GB2312')
        # html = f.read()
        soup = BeautifulSoup(html, "html.parser")

        

        title = soup.findAll(name="a", attrs={"name":"itemlist-title"})
        author = []
        ps = soup.findAll(name='p',attrs={"class":"search_book_author"})
        for p in ps:
            author.append(p.a.string)
        # author = soup.findAll(name="a", attrs={"name":"itemlist-author"})
        pre_price = soup.findAll(name="span", attrs={'class':'search_pre_price'})
        now_price = soup.findAll(name="span", attrs={'class':'search_now_price'})
        publisher = soup.findAll(name="a", attrs={'name':'P_cbs'})

        message = list(zip(title,author,pre_price,now_price,publisher))

        ISBN = [] # 国际索书号
        classies = [] # 书籍类别
        idNumbers = [] # 图书id
        loopStop = 0
        for k in list(title):
            loopStop+=1
            if(loopStop>=20):
                loopStop = 0
                time.sleep(0.5)
            icount = icount + 1
            
            childurl = k.get('href')
            idNumbers.append(findId(childurl))
            ff = urllib.request.urlopen(childurl)
            
            bookImformationHtml = ff.read()
            bookSoup = BeautifulSoup(bookImformationHtml, "html.parser")

            li = bookSoup.find(name="ul", attrs={"class":"key clearfix"})
            tmp = findISBN(str(li))
            ISBN.append(tmp)

            classify = bookSoup.findAll(name = "a",attrs={"class":"green"})
            tmpClass = []  # 临时用于保存书籍类别的list
            for a_ in classify:
                if(a_.string != "图书"):
                    tmpClass.append(a_.string)
                    tmpClass.append(".")
            classies.append(tmpClass[:-1])

            os.system('cls')
            print("关键词：",keyword)
            print("正在获取 : ",icount,"/ ", ranges*len(title))
            print("获取链接: " + childurl)
            print("ID: " + findId(childurl))
            print("书名: " + k['title'])
            print("国际索书号ISBN: " + tmp)
            print("分类: " + ("".join(tmpClass[:-1])).replace('.','>'))

        allMessage = list(zip(title,author,pre_price,now_price,publisher, ISBN, classies, idNumbers))

        iicount = 1
        for item in allMessage:
            if(item[5].isdigit()):
                sheet.write(count,0, item[7]) # id
                sheet.write(count,1, item[0]['title']) # title
                sheet.write(count,2, item[5]) # ISBN
                sheet.write(count,3, item[1]) # Author
                sheet.write(count,4, item[2].string) # pre_price
                sheet.write(count,5, item[3].string) # now_price
                sheet.write(count,6, item[4].string) # publisher
                sheet.write(count,7, item[6]) # classifies
                count = count + 1
            
            os.system('cls')
            print("关键词：",keyword)
            print("正在写入缓存 : ",iicount," / ", len(allMessage))
            iicount = iicount + 1
        excel.save(filename) # xlwt对象的保存方法，这时便覆盖掉了原来的excel
        time.sleep(1) # 暂停 1 秒


def turn2XLSX(filename):
    fname = "D:\\Alask\\Desktop\\ExperimentTwo\\"+filename
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(fname)

    wb.SaveAs(fname+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
    wb.Close()                               #FileFormat = 56 is for .xls extension
    excel.Application.Quit()

def start(keyword, ranges = 1,flag = False):
    for k in keyword:
        kore(k, ranges,flag)
    os.system('cls')
    print("写入完成，正在转换文件格式")
    turnFileName = filename+'x'
    if os.path.exists(turnFileName): # 删除现有文件
        os.remove(turnFileName)
    turn2XLSX(filename) # 转换文件格式
    if(flag == False): deleteSheet(turnFileName," ") # 删除空sheet

def deleteSheet(sExcelFile, sheet): 
    '''
    删除开始时候建的空sheet
    '''
    wb = openpyxl.load_workbook(sExcelFile)
    ws = wb[sheet]
    wb.remove(ws)
    wb.save(sExcelFile)
    print("文件转换完成")



def main():
    keyword = [r"东野圭吾",r"python",r"乙一",r"心里",r"陆秋槎",r"信息系统设计"] # 关键词列表
    start(keyword,2 ,True) # bool：是否只创建一个表;int: 页面数
    os.system('cls')
    print("已完成")

if __name__=="__main__":
    main()

