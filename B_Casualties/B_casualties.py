#!/usr/bin/python
# -*- coding: utf-8 -*-
from xlrd import open_workbook,cellname
import csv
import sys 

#This prints cell values on terminal
book = open_workbook('B_Casualties.xls')
sheet = book.sheet_by_index(book.nsheets-1)
print sheet.name
print sheet.nrows
print sheet.ncols
for row_index in range(sheet.nrows):
    for col_index in range(sheet.ncols):
        print cellname(row_index,col_index),'-',
        print sheet.cell(row_index,col_index).value

#This writes the original xls file (first sheet for index 0) into a csv file.
sheet = book.sheet_by_index(book.nsheets-1)
fp = open(('B_Casualties2_test.csv'), 'wb')
wr = csv.writer(fp, quoting=csv.QUOTE_ALL)
for rownum in xrange(sheet.nrows):
     wr.writerow([unicode(val).encode('utf8') for val in sheet.row_values(rownum)])
fp.close()


#This rewrites a new csv file formatted according to proposed geocoded shapefile join table
with open('B_Casualties_Shp.csv', 'wb') as f:
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

#This rewrites the csv to fit the Data template required by the web2py system
"""#This writes the headers for Incidents Monitored
with open('B_Casualties_Data.csv', 'wb') as f:
    writer = csv.writer(f)
    writer.writerow( ('Template', 'Series', 'Organisation', 'STD-WHO', 'STD-L0', 'STD-L1', 'STD-L2', 'STD-L3', 'STD-Lon', 'STD-Lat', 'STD-DATE', 'STD-TIME', 'IM1', 'IM2', 'IM3', 'IM4') )
    for rownum in xrange(sheet.nrows):
        writer.writerow( ('Incidents Monitored', sheet.cell(2,1).value, sheet.cell(0,3).value, 'Calamity Name Here','Philippines', 'STD-L1', 'STD-L2', sheet.cell(rownum,1).value, 'STD-Lon', 'STD-Lat', sheet.cell(4,1).value, sheet.cell(4,2).value, sheet.cell(rownum,3).value, sheet.cell(rownum,4).value, 1) )
f.close()"""

"""# row 3, column 2
data[2][1] = '20.6'

writer = csv.writer(open('generation.csv', 'wb'))
writer.writerows(data)"""

