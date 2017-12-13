# -*- coding: utf-8 -*-
import os
import xlrd
import xlwt
import collections

filedir="D://a/"

files_list = os.listdir(filedir)

fh = xlrd.open_workbook(filedir+files_list[0])
wb = xlwt.Workbook()
ws = wb.add_sheet("sheet1")
table = fh.sheets()[3]

datavalue = []

for file in files_list:
    fh = xlrd.open_workbook(filedir + file)
    table = fh.sheets()[3]
    nrow = table.nrows
    for i in range(0,nrow):
        data = table.row_values(i)[12]
        if "DMP" in data:
            datavalue.append(data)

d = collections.Counter(datavalue)

n = 0
for k in d :
    ws.write(n,0,k)
    ws.write(n,1,d[k])
    n+=1

wb.save("D://test.xlsx")








