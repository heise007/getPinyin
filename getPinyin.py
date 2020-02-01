#! /usr/bin/env python
#coding:utf-8
#encoding=utf-8
import os
import time
import sys
reload(sys)  
sys.setdefaultencoding('utf8')

import xlrd
import xlwt
from xpinyin import Pinyin

if __name__ == '__main__':
    # global logger
    pingyin=[]
    curr_path = os.getcwd()+'/'
    data= xlrd.open_workbook('./1.xlsx') 
    table = data.sheets()[0]
    nrows = table.nrows
    print nrows
    p=Pinyin()
    

    for index in range(1,nrows):
        #print index
        #print nrows
        #print table.nrows
        #print table.row(index)
        print p.get_pinyin(unicode(str(table.row(index)[2].value))).replace(" ","").replace("-","")
        pingyin.append(p.get_pinyin(unicode(str(table.row(index)[2].value))).replace(" ","").replace("-",""))
        #print p.get_pinyin(str(table.row(index)[2].value))
        #print type(str(table.row(index)[2].value))
        #print type(int(table.row(index)[3].value))

    filename="./pingyin.txt"
    with open(filename, 'w') as file_object:
        for index1 in range(0,len(pingyin)):
            file_object.write(pingyin[index1]+"\r\n")
    workbook = xlwt.Workbook(encoding = 'ascii')
    worksheet = workbook.add_sheet('sheet1')
    worksheet.write(0, 0, label = 'test')
    workbook.save('Myxls.xls')
