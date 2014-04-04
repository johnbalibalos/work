#!/usr/bin/python
import decimal
import xlrd
from xlwt import *
import re
import os
from xlwt.Utils import MAX_ROW, MAX_COL
import csv
import xlwt
from xlrd import cellname, cellnameabs, colname
import sys

#opens database file
f = open('082812annotation.txt', 'rU')
annofile = csv.reader(f, delimiter="\t")
anno = list()
for items in annofile:
    anno.append(items)
#print anno[1]
#print len(anno[1])
#print len(anno[1]) + 1

#excel styles
h_style = easyxf('font: underline single, color blue;')
t_style = easyxf('align: horiz center; font: bold 1, color black; pattern: pattern solid, fore_color yellow;')
t2_style = easyxf('align: horiz center; font: bold 1, color black; pattern: pattern solid, fore_color yellow; borders: bottom medium;')
error_style = easyxf('pattern: pattern solid, fore_color red;')
#sum_style = easyxf('align: horiz center; font: bold 1, color black; pattern: pattern solid, fore_color yellow; borders: bottom medium;')
decimal_style = easyxf("", "#,###.000")
sc_style = easyxf('pattern: pattern solid, fore_color lime')
sc0_style = easyxf('font: color 22;')
title_style = easyxf('align: horiz right;')
avgSC_style = easyxf("", "#,###")

#protein list input
inputfile = raw_input("Type in the full path for sample report file:")

#Read and create excel workbook and worksheets
readbook = xlrd.open_workbook(inputfile, formatting_info=True)
rs = readbook.sheet_by_index(0)
writebook = xlwt.Workbook(encoding='utf-8')
writebook.add_sheet('errors')
writebook.add_sheet('duplicate')
writebook.add_sheet('original')
writebook.add_sheet('Annotation')
writebook.add_sheet('NSAF')
duplicate = writebook.get_sheet(1)
error = writebook.get_sheet(0)
original = writebook.get_sheet(2)
Annotation = writebook.get_sheet(3)
Nsaf = writebook.get_sheet(4)

row = 1
col = 1
scaacol = rs.ncols - col
nsafcol = 2 * scaacol
#Sample name
def samplename(sheet,column,title):
    newcol = len(anno[1]) + 1
    col = 1
    for col_index in range(col, rs.ncols):
        sheet.write(0, newcol+column, rs.cell(0,col_index).value, t_style)
        sheet.write(1, newcol+column, "", t_style)
        sheet.write(2, newcol+column, title, t2_style)
        newcol += 1

def checker():
    global Noscaa
    Noscaa = raw_input("Would you like SC/AA information? Yes/No/Quit:")
    if Noscaa == 'Yes':
        if rs.ncols > 71:
            print 'Too many samples for SC/AA information. Please split to 70 samples or fewer.'
            checker()
    elif Noscaa == 'No':
        if rs.ncols > 114:
            print 'Too many samples. Please split file to 110 samples or fewer.' checker() elif Noscaa == 'Quit': sys.exit() else: checker() checker() #Copies original sheet to new workbook row = 0
col = 0
newrow = 0
newcol = 0
for col_index in range(0, rs.ncols):
    for row_index in range(0, rs.nrows):
        original.write(newrow, newcol, rs.cell(row_index, col_index).value)
        newrow += 1
    newrow = 0
    newcol += 1

#Annotation headers
def annoheader(sheet):
    col = 1
    f.seek(0)
    headers = annofile.next()
    for titles in headers:
        sheet.write(0, col, titles, t_style)
        col += 1
    sheet.write(0, 0, "", t_style)
    annocol = len(anno[1]) + 1
    for col_index in range(0, annocol):
        sheet.write(1, col_index, "", t_style)
        sheet.write(2, col_index, "", t2_style)

annoheader(Annotation)
annoheader(Nsaf)

row = 1
newrow = 3
col = 1
newcol = len(anno[1]) + 1
analysis = 3

Nsaf.write(0, newcol + rs.ncols - 1, "", t_style)
Nsaf.write(0, newcol + rs.ncols, "", t_style)
Nsaf.write(1, newcol + rs.ncols - 1, "", t_style)
Nsaf.write(1, newcol + rs.ncols, "", t_style)
Nsaf.write(2, newcol + rs.ncols - 1, 'SC >0', t2_style)
Nsaf.write(2, newcol + rs.ncols, 'SC >10', t2_style)
Nsaf.write(0, newcol + rs.ncols + 1, "", t_style)
Nsaf.write(1, newcol + rs.ncols + 1, "", t_style)
Nsaf.write(2, newcol + rs.ncols + 1, 'AVE SC > 0', t2_style)
if Noscaa == 'No':
    Nsaf.write(0, newcol + nsafcol + analysis, "", t_style)
    Nsaf.write(1, newcol + nsafcol + analysis, "", t_style)
    Nsaf.write(2, newcol + nsafcol + analysis, 'Sum NSAF', t2_style)
    Nsaf.write(0, newcol + nsafcol + analysis + 1, "", t_style)
    Nsaf.write(1, newcol + nsafcol + analysis + 1, "", t_style)
    Nsaf.write(2, newcol + nsafcol + analysis + 1, 'Avg NSAF', t2_style)
if Noscaa == 'Yes':
    Nsaf.write(0, newcol + scaacol + nsafcol + analysis, "", t_style)
    Nsaf.write(1, newcol + scaacol + nsafcol + analysis, "", t_style)
    Nsaf.write(2, newcol + scaacol + nsafcol + analysis, 'Sum NSAF', t2_style)
    Nsaf.write(0, newcol + scaacol + nsafcol + analysis + 1, "", t_style)
    Nsaf.write(1, newcol + scaacol + nsafcol + analysis + 1, "", t_style)
    Nsaf.write(2, newcol + scaacol + nsafcol + analysis + 1, 'Avg NSAF', t2_style)

for row_index in range(row, rs.nrows):
    SCrange = cellnameabs(newrow, newcol) + ':' + cellnameabs(newrow, newcol + rs.ncols - 2)
    NSAFrange = cellnameabs(newrow, newcol+scaacol+analysis) + ':' + cellnameabs(newrow, newcol + nsafcol + analysis - 1)
    countifSC = 'COUNTIF(' + SCrange + ', ">0"' + ')'
    countif10SC = 'COUNTIF(' + SCrange + ', ">10"' + ')'
    avgSC = 'AVERAGE(' + SCrange + ')'
    avgNSAF = 'AVERAGE(' + NSAFrange + ')'
    sumNSAF = 'SUM(' + NSAFrange + ')'
    Nsaf.write(newrow, newcol + rs.ncols - 1, Formula(countifSC))
    Nsaf.write(newrow, newcol + rs.ncols, Formula(countif10SC))
    Nsaf.write(newrow, newcol + rs.ncols + 1, Formula(avgSC), avgSC_style)
    if Noscaa == 'No':
        Nsaf.write(newrow, newcol + nsafcol + analysis, Formula(sumNSAF), decimal_style)
        Nsaf.write(newrow, newcol + nsafcol + analysis + 1, Formula(avgNSAF), decimal_style)
    if Noscaa == 'Yes':
        NSAFrange = cellnameabs(newrow, newcol+nsafcol+analysis) + ':' + cellnameabs(newrow, newcol + scaacol + nsafcol + analysis - 1)
        avgNSAF = 'AVERAGE(' + NSAFrange + ')'
        sumNSAF = 'SUM(' + NSAFrange + ')'
        Nsaf.write(newrow, newcol + scaacol + nsafcol + analysis, Formula(sumNSAF), decimal_style)
        Nsaf.write(newrow, newcol + scaacol + nsafcol + analysis + 1, Formula(avgNSAF), decimal_style)
    newrow += 1

#Hyperlinks and annotation information
proteinlist = list()
noprotein = list()
lengthlist = list()
duplicateprotein = list()
def Hyperlink(sheet):
    row = 1
    col = 1
    newrow = 3
    newcol = 1
    duprow = 0
    for row_index in range(row, rs.nrows):
        proteinfound = False
        protein = rs.cell(row_index, 0).value
        pattern = r'_HUMAN.*'
        replacement = r'_HUMAN'
        outprotein = re.sub(pattern, replacement, protein)
        duplicateprot = False
        if protein != "":
            ##Hyperlinks
            uniprot = 'http://www.uniprot.org/uniprot/'
            link = 'HYPERLINK(' + '"' + (uniprot + outprotein) + '"' + '; "' + protein + '")'
            Annotation.write(newrow, 0, Formula(link), h_style)
            sheet.write(newrow, 0, Formula(link), h_style)
            ##annotation
            p = re.compile(r'(?i)\b(?:%s)\b' % outprotein)
            counter = 0
            for annotation in anno:
                if bool(p.match(annotation[1])) == True:
                    proteinfound = True
                    if counter == 0:
                        counter += 1
                        lengthlist.append(annotation[23])
                        for items in annotation:
                            #print items
                            Annotation.write(newrow, newcol, items)
                            sheet.write(newrow, newcol, items)
                            newcol += 1
                    elif counter > 0:
                        proteinfound = True
                        counter += 1
                        duplicateprot = True
                    numberinlist = row_index
                    totalrows = rs.nrows - row
                    print 'protein %s of %s proteins' % (numberinlist, totalrows)
            for col_index in range(col, rs.ncols): #Spectral count
                SC0 = newrow, newcol, rs.cell(row_index, col_index).value, sc0_style
                SC = newrow, newcol, rs.cell(row_index, col_index).value, sc_style
                if rs.cell(row_index, col_index).value == 0:
                    Annotation.write(*SC0)
                    sheet.write(*SC0)
                elif rs.cell(row_index, col_index).value != 0:
                    Annotation.write(*SC)
                    sheet.write(*SC)
                newcol += 1
        if duplicateprot == True:
            duplicate.write(duprow, 0, protein)
            duprow += 1
        if proteinfound == True:
            newcol = 1
            newrow += 1
        elif proteinfound == False:
            lengthlist.append(0)
            noprotein.append(outprotein)
            newcol = 1
            newrow += 1
            #Prints proteins not found in file
            print '%s not found in annotation file' % outprotein

Hyperlink(Nsaf)

if Noscaa == 'Yes':
    samplename(Annotation,0,'SC')
    samplename(Nsaf,0,'SC')
    samplename(Nsaf,scaacol+analysis,'SC/AA')
    samplename(Nsaf,nsafcol+analysis,'NSAF')
if Noscaa == 'No':
    samplename(Annotation,0,'SC')
    samplename(Nsaf,0,'SC')
    samplename(Nsaf,scaacol+analysis,'NSAF')

##SC/AA information
col = 1
row = 1
newrow = 3
newcol = len(anno[1]) + 1
i = 0
scaa_array = list()
for row_index in range(row, rs.nrows):
    newcol = len(anno[1]) + 1
    scaalist = list()
    #print lengthlist[i], len(lengthlist)
    for col_index in range(col, rs.ncols):
        if lengthlist[i] == 0:
            scaa = 0
            scaalist.append(float(scaa))
            if Noscaa == 'Yes':
                Nsaf.write(newrow,newcol+scaacol+analysis,scaa,error_style)
            newcol += 1
        elif lengthlist[i] == '':
            scaa = 0
            scaalist.append(float(scaa))
            if Noscaa == 'Yes':
                Nsaf.write(newrow,newcol+scaacol+analysis,scaa,error_style)
            newcol += 1
        else:
            scaa = decimal.Decimal(str(rs.cell(row_index,col_index).value)) / decimal.Decimal(str(lengthlist[i]))
            if Noscaa == 'Yes':
                Nsaf.write(newrow,newcol+scaacol+analysis,scaa,decimal_style)
            scaalist.append(float(scaa))
            #print scaalist
            newcol += 1
    i += 1
    newrow += 1
    scaa_array.append(scaalist)
#scaa_array is a list of lists containing: sc/aa values in rows for list[i] of scaa_array
#print scaa_array

##NSAF information
sumscaalist = list()
i = 0
for y in range(len(scaa_array[i])):
    sum = 0
    for x in range(len(scaa_array)):
        sum += scaa_array[x][y]
    sumscaalist.append(sum)

newrow = rs.nrows + 3
newcol = len(anno[1]) + 1
for y in range(col, rs.ncols):
    if Noscaa == 'Yes':
        Nsaf.write(newrow-1,newcol+scaacol+analysis,'Sum SC/AA',t_style)
    #Nsaf.write(newrow-1,newcol+nsafcol,'Sum NSAF',t_style)
    newcol += 1
col = 1
row = 1
newcol = len(anno[1]) + 1
for num in sumscaalist:
    if Noscaa == 'Yes':
        Nsaf.write(newrow,newcol+scaacol+analysis,num,decimal_style)
    newcol += 1

newcol = len(anno[1]) + 1
newrow = 3
j = 0
nsafcol = 2*scaacol
for num in sumscaalist:
    newrow = 3
    for i in range(len(scaa_array)):
        NSAF = decimal.Decimal(str(100* scaa_array[i][j])) / decimal.Decimal(str(num))
        if Noscaa == 'Yes':
            Nsaf.write(newrow, newcol+nsafcol+analysis,NSAF,decimal_style)
        if Noscaa == 'No':
            Nsaf.write(newrow, newcol+scaacol+analysis,NSAF,decimal_style)
        newrow += 1
    j += 1
    newcol += 1
#Prints list of proteins not found
error_row = 0
for items in noprotein:
    print '%s not found in annotation file' % items
    error.write(error_row, 0, items)
    error_row += 1
length = len(noprotein)
print '%s proteins not found' % length

#Saves new workbook to .out file
savefile = os.path.splitext(inputfile)[0] + u'.out' + os.path.splitext(inputfile)[-1]
print 'Job is done. Your file is located here:' + savefile
writebook.save(savefile)
f.close()
