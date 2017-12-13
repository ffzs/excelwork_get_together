# -*- coding: utf-8 -*-
import os
import xlrd
import xlwt

filedir="F://exceldir/"

files_list = os.listdir(filedir)

fh = xlrd.open_workbook(filedir+files_list[0])
sheet_names = fh.sheet_names()
sheet_num = len(sheet_names)
wb = xlwt.Workbook()

for num in range(0,sheet_num):
    sheetname = sheet_names[num]
    ws = wb.add_sheet(sheetname)
    fh = xlrd.open_workbook(filedir+files_list[0])
    table = fh.sheets()[num]

    datavalue = []
    datavalue.append(table.row_values(0))

    for file in files_list:
        fh = xlrd.open_workbook(filedir + file)
        table = fh.sheets()[num]
        nrow = table.nrows
        for i in range(1,nrow):
            rdata = table.row_values(i)
            if rdata[30]=="微生物建库组":
                datavalue.append(rdata)

    for a in range(len(datavalue)):
        for b in range(len(datavalue[a])):
            c = datavalue[a][b]
            ws.write(a,b,c)

wb.save("F://合并.xls")








