#!/usr/bin/python
import xlrd
from xlwt import *
import re
import os
from xlwt.Utils import MAX_ROW, MAX_COL
import csv
import xlwt
from xlrd import cellname, cellnameabs, colname
import readline

#writing functions
def iterwrite(x, y, new_x, new_y, worksheet, modulator):
    for y_index in range(y, rs.ncols):
        for x_index in range(x, rs.nrows):
            cell_value = rs.cell(x_index, y_index).value
            print cell_value
            
            new_x += 1
        new_x = 0
    new_y += 1

#opens database file
f = open('sp2_metadata_biodb.txt', 'rb')
annofile = csv.reader(f, delimiter="\t")
anno = list()
for items in annofile:
    anno.append(items)

#excel styles
h_style = easyxf('font: underline single, color blue;')
t_style = easyxf('align: horiz center; font: bold 1, color black; pattern: pattern solid, fore_color yellow;')
t2_style = easyxf('align: horiz center; font: bold 1, color black; pattern: pattern solid, fore_color yellow; borders: bottom medium;')
error_style = easyxf('font: bold 1, color red;')
decimal_style = easyxf("", "#,###.000")
sc_style = easyxf('pattern: pattern solid, fore_color sea_green;')
sc0_style = easyxf('font: color 22;')
title_style = easyxf('align: horiz right;')

#protein list input
inputfile = raw_input("Type in the full path for sample report file:")

#Read and create excel workbook and worksheets
Readbook = xlrd.open_workbook(inputfile, formatting_info=True)
rs = Readbook.sheet_by_index(0)
Writebook = xlwt.Workbook(encoding='utf-8')
Writebook.add_sheet('original')
Writebook.add_sheet('match')
Writebook.add_sheet('NSAF')
wo = Writebook.get_sheet(0)
wm = Writebook.get_sheet(1)
wn = Writebook.get_sheet(2)

#Duplicate original sheet
iterwrite(0,0,0,0, wo, 

savefile = os.path.splitext(inputfile)[0] + u'.out' + os.path.splitext(inputfile)[-1]
print 'Job is done. Your file is located here:' + savefile
Writebook.save(savefile)
f.close()


