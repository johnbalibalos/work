#!/usr/bin/python
from xlutils.copy import copy
import xlrd
from xlwt import * 
import re
import os
from xlwt.Utils import MAX_ROW, MAX_COL
import csv
import xlwt
from xlrd import cellname, cellnameabs, colname

#opens database file
f = open('sp2_metadata_biodb.txt', 'rb')
annofile = csv.reader(f, delimiter="\t")

h_style = easyxf('font: underline single, color blue;')
t_style = easyxf('align: horiz center; font: bold 1, color black; pattern: pattern solid, fore_color yellow;')
error_style = easyxf('font: bold 1, color red;')
decimal_style = easyxf("", "#,###.000")

#protein list input from file
inputfile = raw_input("Type in the full path for sample report file:")

rb = xlrd.open_workbook(inputfile,formatting_info=True)
#change index to (0) when using original sample report
rs = rb.sheet_by_index(0)
wb = xlwt.Workbook(encoding='utf-8')
wb.add_sheet('original')
wb.add_sheet('match')
wb.add_sheet('NSAF')
wo = wb.get_sheet(0)
ws = wb.get_sheet(1)
wn = wb.get_sheet(2)

row = 0
col = 0
newrow = 0
newcol = 0
for col_index in range(col, rs.ncols):
	for row_index in range(row, rs.nrows):
		wo.write(newrow, newcol, rs.cell(row_index, col_index).value)
		newrow += 1
	newrow = 0
	newcol += 1

proteinlist = list()

row = 29
newrow = 0
samples = 9
newsamples = 4

ws.write(newrow, 0, rs.cell(row, 0).value, t_style)
ws.write(newrow, 1, rs.cell(row, 3).value, t_style)
ws.write(newrow, 2, rs.cell(row, 4).value, t_style)
ws.write(newrow, 3, rs.cell(row, 5).value, t_style)
wn.write(newrow, 0, rs.cell(row, 0).value, t_style)
wn.write(newrow, 1, rs.cell(row, 3).value, t_style)
wn.write(newrow, 2, rs.cell(row, 4).value, t_style)
wn.write(newrow, 3, rs.cell(row, 5).value, t_style)

#makes row 2 yellow
for col_index in range(0, 26):
	ws.write(1, col_index, "", t_style)
	wn.write(1, col_index, "", t_style)

#title of columns
def titles():
	ws.write(newrow, 0, rs.cell(row_index, 0).value)
	ws.write(newrow, 1, rs.cell(row_index, 3).value)
	ws.write(newrow, 3, rs.cell(row_index, 5).value)
	wn.write(newrow, 0, rs.cell(row_index, 0).value)
	wn.write(newrow, 1, rs.cell(row_index, 3).value)
	wn.write(newrow, 3, rs.cell(row_index, 5).value)

#iterates through rows and copies cell value hyperlink to new worksheet 
row = 30
newrow = 2
for row_index in range(row, rs.nrows):
	protein = rs.cell(row_index,4).value
	pattern = r'_HUMAN.*'
	replacement = '_HUMAN'
	outprotein = re.sub(pattern, replacement, protein)
	if bool(re.search("Accession.*", rs.cell(row_index, 4).value)) == True:
		ws.write(newrow, 2, rs.cell(row_index, 4).value)
		wn.write(newrow, 2, rs.cell(row_index, 4).value)
		titles()
	elif outprotein != "":
		uniprot = 'http://www.uniprot.org/uniprot/'
		link = 'HYPERLINK(' + '"' + (uniprot + outprotein) + '"' + '; "' + protein + '")'
		ws.write(newrow, 2, Formula(link), h_style)
		wn.write(newrow, 2, Formula(link), h_style)
		titles()
	newrow += 1
#headers for annotation information
headers = annofile.next()
column = 4
for titles in headers:
	ws.write(0, column, titles, t_style)
	wn.write(0, column, titles, t_style)
	column +=1

#annotation information
row = 30
newrow = 2
newcol = 4
noprotein = list()
anno = list()
col = 9
scaacol = (rs.ncols - col)
nsafcol = (2*(rs.ncols - col))


for items in annofile:
	anno.append(items)
#print anno[1]

for row_index in range(row, rs.nrows):
	proteinfound = False
	f.seek(0)
	protein = rs.cell(row_index, 4).value
	pattern = r'_HUMAN.*'
	replacement = r'_HUMAN'
	outprotein = re.sub(pattern, replacement, protein)
	for annotation in anno:
		if protein != "":
			#print annotation
			#print annotation[1]
			p = re.compile(r'(?i)\b(?:%s)\b' % outprotein)
			if bool(p.match(annotation[1])) == True:
				for items in annotation:
					ws.write(newrow, newcol, items)
					wn.write(newrow, newcol, items)
					newcol += 1
					proteinfound = True
				#SC/AA and NSAF information
				newcol = 26
				aalength = cellnameabs(newrow, 7)
				col = 9
				scaalist = list()
				noscaa = list()
				for col_i in range(col, rs.ncols):
					try:
						#scaa = float(rs.cell(row_index, col_i).value) / float(aalength)
						#wn.write(newrow, (newcol + scaacol), scaa)
						if annotation[3] == '-':
							wn.write(newrow, (newcol + scaacol), 0, error_style)
							newcol += 1
						else:	
							wn.write(newrow, (newcol + scaacol), Formula(cellnameabs(newrow,newcol) + '/' + aalength),decimal_style)
							wn.write(newrow, (newcol + nsafcol), Formula('(100*' + cellnameabs(newrow,newcol+scaacol) + ')/' + cellnameabs(rs.nrows-29, newcol+scaacol)),decimal_style)
							newcol += 1
					except:
						#wn.write(newrow, (newcol + scaacol), Formula(cellnameabs(newrow,newcol) + '/' + aalength))
						if annotation[3] == '-':
							wn.write(newrow, (newcol + scaacol), 0, error_style)
							newcol += 1
						else:
							wn.write(newrow, (newcol + scaacol), 0, error_style)
							wn.write(newrow, (newcol + nsafcol), Formula('(100*' + cellnameabs(newrow,newcol+scaacol) + ')/' + cellnameabs(rs.nrows-29, newcol+scaacol)),decimal_style)
							newcol += 1
				###########################
				newcol = 4
				newrow += 1
				numberinlist = row_index - row + 1
				totalrows = rs.nrows - row
				print 'protein %s of %s proteins' % (numberinlist, totalrows)


	if proteinfound == False:
		noprotein.append(outprotein)
		newcol = 4
		newrow += 1
		#Prints proteins that are not found in file
		print '%s not found in annotation file' % outprotein

#print noprotein and noscaa
for items in noprotein:	
	print '%s not found in annotation file' % items
length = len(noprotein)
print '%s proteins not found' % length

#iterates through columns and rows to copy spectral count data
row = 29
newrow = 0
col = 9
newcol = 26
for col_index in range(col, rs.ncols):
	sumscaa = 'SUM(' + cellnameabs(2,newcol + scaacol) + ':' + cellnameabs((rs.nrows - 31),newcol + scaacol) + ')'
	sumnsaf = 'SUM(' + cellnameabs(2,newcol + nsafcol) + ':' + cellnameabs((rs.nrows - 31),newcol + nsafcol) + ')'
	ws.write(newrow, newcol, rs.cell(row, col_index).value, t_style)
	wn.write(newrow, newcol, rs.cell(row, col_index).value, t_style)
	wn.write(newrow, (newcol + (rs.ncols - col)), rs.cell(row, col_index).value, t_style)
	wn.write(newrow, (newcol + (2*(rs.ncols - col))), rs.cell(row, col_index).value, t_style)
	wn.write(1, newcol, 'SC', t_style)
	ws.write(1, newcol, 'SC', t_style)
	wn.write(1, (newcol + (rs.ncols - col)), 'SC/AA', t_style)
	wn.write(1, (newcol + (2*(rs.ncols - col))), '%NSAF', t_style)
	wn.write((rs.nrows - 28), newcol + scaacol, 'SUM SC/AA')
	wn.write((rs.nrows - 28), newcol + nsafcol, 'SUM NSAF%')
	wn.write((rs.nrows - 29), newcol + scaacol, Formula(sumscaa),decimal_style)
	wn.write((rs.nrows - 29), newcol + nsafcol, Formula(sumnsaf),decimal_style)
	newcol += 1
row = 30
newrow = 2 
col = 9
newcol = 26

#Spectral count numbers
for col_index in range(col, rs.ncols):
	for row_index in range(row, rs.nrows):
		if rs.cell(row_index, col).value != "":
			ws.write(newrow, newcol, rs.cell(row_index, col).value)
			wn.write(newrow, newcol, rs.cell(row_index, col).value)
			newrow += 1
	newrow = 2
	col += 1
	newcol += 1


#save output to excel
savefile = os.path.splitext(inputfile)[0] + u'.out' + os.path.splitext(inputfile)[-1]
print 'Job is done. Your file is located here:' + savefile
wb.save(savefile)
f.close()
