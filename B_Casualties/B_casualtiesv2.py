#!/usr/bin/python
# -*- coding: utf-8 -*-
from xlrd import open_workbook,cellname
from xlutils.copy import copy # http://pypi.python.org/pypi/xlutils
from xlwt import easyxf # http://pypi.python.org/pypi/xlwt
import sys
import csv


rb = open_workbook('B_Casualties.xls',formatting_info=True)
r_sheet = rb.sheet_by_index(0) # read only copy to introspect the file
wb = copy(rb) # a writable copy (I can't read values out of this, only write to it)
w_sheet = wb.get_sheet(0) # the sheet to write to within the writable copy

#Transfer DEAD MISSING AND INJURED LABELS ELSEWHERE
for row_index in xrange(r_sheet.nrows):
    casualtytype = r_sheet.cell(row_index, 3).value.encode('ascii', 'ignore').upper()
    if casualtytype == "DEAD":
        w_sheet.write(row_index, r_sheet.ncols, "DEAD")
    elif casualtytype == "INJURED":
        w_sheet.write(row_index, r_sheet.ncols, "INJURED")
    elif casualtytype == "MISSING":
        w_sheet.write(row_index, r_sheet.ncols, "MISSING")
    elif len(r_sheet.cell(row_index,3).value)>7:
        w_sheet.write(row_index, r_sheet.ncols, "nan") 

wb.save('B_Casualties.xls')

rb = open_workbook('B_Casualties.xls',formatting_info=True)
r_sheet = rb.sheet_by_index(0) # read only copy to introspect the file
wb = copy(rb) # a writable copy 
w_sheet = wb.get_sheet(0) # the sheet to write to within the writable copy
counter=0


#Tag each row with corresponding dead, missing, injured
for row_index in xrange(r_sheet.nrows):
    if len(r_sheet.cell(row_index,3).value)>7:
        while len(r_sheet.cell(row_index-counter,r_sheet.ncols-1).value)<4:
            counter=counter+1
        w_sheet.write(row_index, r_sheet.ncols-1, r_sheet.cell(row_index-counter,r_sheet.ncols-1).value)
        counter=0

#Save new file
wb.save('B_Casualties.xls')


#This prints cell values on terminal
book = open_workbook('B_Casualties.xls')
sheet = book.sheet_by_index(0)
print sheet.name
print sheet.nrows
print sheet.ncols
for row_index in range(sheet.nrows):
    for col_index in range(sheet.ncols):
        print cellname(row_index,col_index),'-',
        print sheet.cell(row_index,col_index).value

#This writes the original xls file (first sheet for index 0) into a csv file.
sheet = book.sheet_by_index(0)
fp = open(('B_Casualties.csv'), 'wb')
wr = csv.writer(fp, quoting=csv.QUOTE_ALL)
for rownum in xrange(sheet.nrows):
     wr.writerow([unicode(val).encode('utf8') for val in sheet.row_values(rownum)])
fp.close()


#This rewrites a new csv file formatted according to proposed geocoded shapefile join table
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
f.close()
