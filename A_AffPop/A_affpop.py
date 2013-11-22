#!/usr/bin/python
# -*- coding: utf-8 -*-
from xlrd import open_workbook,cellname
from xlutils.copy import copy # http://pypi.python.org/pypi/xlutils
from xlwt import easyxf # http://pypi.python.org/pypi/xlwt
import sys
import csv

#This prints cell values on terminal
book = open_workbook('AffPop_input.xls')
sheet = book.sheet_by_index(10)
print sheet.name
print sheet.nrows
print sheet.ncols
for row_index in range(sheet.nrows):
    for col_index in range(sheet.ncols):
        print cellname(row_index,col_index),'-',
        print sheet.cell(row_index,col_index).value

#This writes the original xls file (first sheet for index 0) into a csv file.
sheet = book.sheet_by_index(0)
fp = open(('AffPop_raw.csv'), 'wb')
wr = csv.writer(fp, quoting=csv.QUOTE_ALL)
for rownum in xrange(sheet.nrows):
     wr.writerow([unicode(val).encode('utf8') for val in sheet.row_values(rownum)])
fp.close()


"""#This rewrites a new csv file formatted according to proposed geocoded shapefile join table
with open('B_Casualties.csv', 'wb') as f:
    writer = csv.writer(f)
    writer.writerow( ('SR_NO', 'TYPE', 'PROVINCE', 'NAME', 'AGE', 'ADDRESS', 'REMARKS', 'DATETIME', 'DEAD', 'INJURED', 'MISSING', 'TAGCHECK') )
    calamityname = raw_input("What type of calamity is this?")
    counter = 1
    for rownum in xrange(sheet.nrows):
            #This conditional will eliminate NAME, MISSING, DEAD, INJURED rows from main entries
            if len(sheet.cell(rownum,3).value)>7 & len(sheet.cell(rownum, sheet.ncols-1).value)==4:
                if len(sheet.cell(rownum,1).value)>0:
                    writer.writerow( (sheet.cell(0,1).value.encode('ascii', 'ignore').upper(), calamityname.upper(), sheet.cell(rownum,1).value.encode('ascii', 'ignore').upper(), sheet.cell(rownum,3).value.encode('ascii', 'ignore').upper(), sheet.cell(rownum,5).value.encode('ascii', 'ignore').upper(), sheet.cell(rownum,6).value.encode('ascii', 'ignore').upper(), sheet.cell(rownum,7).value.encode('ascii', 'ignore').upper(), sheet.cell(3,1).value.encode('ascii', 'ignore').upper(), sheet.cell(rownum,2).value, 0, 0, sheet.cell(rownum, sheet.ncols-1).value.encode('ascii', 'ignore').upper() ) )
                else:
                #This will repeat the province name for empty cells
                    while len(sheet.cell(rownum-counter,1).value)==0:
                        counter = counter+1
                    writer.writerow( (sheet.cell(0,1).value.encode('ascii', 'ignore').upper(), calamityname.upper(), sheet.cell((rownum-counter),1).value.encode('ascii', 'ignore').upper(), sheet.cell(rownum,3).value.encode('ascii', 'ignore').upper(), sheet.cell(rownum,5).value.encode('ascii', 'ignore').upper(), sheet.cell(rownum,6).value.encode('ascii', 'ignore').upper(), sheet.cell(rownum,7).value.encode('ascii', 'ignore').upper(), sheet.cell(3,1).value.encode('ascii', 'ignore').upper(), sheet.cell(rownum,2).value, 0, 0, sheet.cell(rownum, sheet.ncols-1).value.encode('ascii', 'ignore').upper() ) )
                    counter = 1
            elif len(sheet.cell(rownum,3).value)>7 & len(sheet.cell(rownum, sheet.ncols-1).value)==7:
                if sheet.cell(rownum, sheet.ncols-1).value=="INJURED":
                    if len(sheet.cell(rownum,1).value)>0:
                        writer.writerow( (sheet.cell(0,1).value.encode('ascii', 'ignore').upper(), calamityname.upper(), sheet.cell(rownum,1).value.encode('ascii', 'ignore').upper(), sheet.cell(rownum,3).value.encode('ascii', 'ignore').upper(), sheet.cell(rownum,5).value.encode('ascii', 'ignore').upper(), sheet.cell(rownum,6).value.encode('ascii', 'ignore').upper(), sheet.cell(rownum,7).value.encode('ascii', 'ignore').upper(), sheet.cell(3,1).value.encode('ascii', 'ignore').upper(), 0, sheet.cell(rownum,2).value, 0, sheet.cell(rownum, sheet.ncols-1).value.encode('ascii', 'ignore').upper() ) )
                    else:
                    #This will repeat the province name for empty cells
                        while len(sheet.cell(rownum-counter,1).value)==0:
                            counter = counter+1
                        writer.writerow( (sheet.cell(0,1).value.encode('ascii', 'ignore').upper(), calamityname.upper(), sheet.cell((rownum-counter),1).value.encode('ascii', 'ignore').upper(), sheet.cell(rownum,3).value.encode('ascii', 'ignore').upper(), sheet.cell(rownum,5).value.encode('ascii', 'ignore').upper(), sheet.cell(rownum,6).value.encode('ascii', 'ignore').upper(), sheet.cell(rownum,7).value.encode('ascii', 'ignore').upper(), sheet.cell(3,1).value.encode('ascii', 'ignore').upper(), 0, sheet.cell(rownum,2).value, 0, sheet.cell(rownum, sheet.ncols-1).value.encode('ascii', 'ignore').upper() ) )
                        counter = 1
                elif sheet.cell(rownum, sheet.ncols-1).value=="MISSING":
                    if len(sheet.cell(rownum,1).value)>0:
                        writer.writerow( (sheet.cell(0,1).value.encode('ascii', 'ignore').upper(), calamityname.upper(), sheet.cell(rownum,1).value.encode('ascii', 'ignore').upper(), sheet.cell(rownum,3).value.encode('ascii', 'ignore').upper(), sheet.cell(rownum,5).value.encode('ascii', 'ignore').upper(), sheet.cell(rownum,6).value.encode('ascii', 'ignore').upper(), sheet.cell(rownum,7).value.encode('ascii', 'ignore').upper(), sheet.cell(3,1).value.encode('ascii', 'ignore').upper(), 0, 0, sheet.cell(rownum,2).value, sheet.cell(rownum, sheet.ncols-1).value.encode('ascii', 'ignore').upper() ) )
                    else:
                    #This will repeat the province name for empty cells
                        while len(sheet.cell(rownum-counter,1).value)==0:
                            counter = counter+1
                        writer.writerow( (sheet.cell(0,1).value.encode('ascii', 'ignore').upper(), calamityname.upper(), sheet.cell((rownum-counter),1).value.encode('ascii', 'ignore').upper(), sheet.cell(rownum,3).value.encode('ascii', 'ignore').upper(), sheet.cell(rownum,5).value.encode('ascii', 'ignore').upper(), sheet.cell(rownum,6).value.encode('ascii', 'ignore').upper(), sheet.cell(rownum,7).value.encode('ascii', 'ignore').upper(), sheet.cell(3,1).value.encode('ascii', 'ignore').upper(), 0, 0, sheet.cell(rownum,2).value, sheet.cell(rownum, sheet.ncols-1).value.encode('ascii', 'ignore').upper() ) )
                        counter = 1
f.close()"""

