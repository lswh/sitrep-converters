#!/usr/bin/python
# -*- coding: utf-8 -*-
from xlrd import open_workbook,cellname
from xlutils.copy import copy # http://pypi.python.org/pypi/xlutils
from xlwt import easyxf # http://pypi.python.org/pypi/xlwt
import sys
import csv
import time

#This prints cell values on terminal
book = open_workbook('AffPop_input.xls')
sheet = book.sheet_by_index(book.nsheets-1)
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


#VANE
#This prints cell values on terminal
book = open_workbook('DAMAGED HOUSES.xls', formatting_info=True)
sheet = book.sheet_by_index(book.nsheets-1)
print book.nsheets
print sheet.name
print sheet.nrows
print sheet.ncols
print sheet.cell(22,2).xf_index
fmt = book.xf_list[sheet.cell(21,2).xf_index]
print fmt.alignment.hor_align
##for row_index in range(sheet.nrows):
##    for col_index in range(sheet.ncols):
##        print cellname(row_index,col_index),'-',
##        print sheet.cell(row_index,col_index).value

#This writes the original xls file (first sheet for index 0) into a csv file.
##sheet = book.sheet_by_index(6)
##fp = open(('DAMAGED_HOUSES.csv'), 'wb')
##wr = csv.writer(fp, quoting=csv.QUOTE_ALL)
####wr = csv.writer(fp, delimiter=',', quoting=csv.QUOTE_NONE)
##for rownum in xrange(sheet.nrows):
##     wr.writerow([unicode(val).encode('utf8') for val in sheet.row_values(rownum)])
##fp.close()


#This rewrites a new csv file formatted according to proposed geocoded shapefile join table
with open('Affected_Population_SHP.csv', 'wb') as f:
     writer = csv.writer(f, quoting=csv.QUOTE_ALL)
     writer.writerow( ('SR_NO', 'TYPE', 'REGION', 'PROVINCE', 'MUNICIPALITY', 'PROVGEOCODE', 'MUNIGEOCODE', 'AFFBGYS', 'AFFFAMILIES', 'AFFPERSONS', 'EVACCTRS', 'IEC_FAM', 'IEC_PERSONS', 'OEC_FAM', 'OEC_PERSONS', 'SERVED_FAMS', 'SERVED_PERSONS') )
     calamityname = raw_input("What type of calamity is this?")
    
     regCounter = 0
     provCounter = 1
     for rownum in xrange(12, sheet.nrows):
          region = 'region'
          province = 'province'
          muniAlign = book.xf_list[sheet.cell(rownum,2).xf_index].alignment.hor_align

          #Selects only municipality entries based on cell alignment
          if len(sheet.cell(rownum, 2).value) > 0 and 'region' not in str(sheet.cell(rownum,1).value.encode('ascii', 'ignore')).lower() and muniAlign == 3:
               #Repeats 'region' values for empty cells
               while len(sheet.cell(rownum-regCounter,1).value)==0:
                    regCounter = regCounter +1
               region = str(sheet.cell(rownum-regCounter, 1).value).upper().replace('REGION','').strip()

               #Repeats 'province' values for empty cells
               while book.xf_list[sheet.cell(rownum-provCounter,2).xf_index].alignment.hor_align != 1:
                    provCounter = provCounter + 1
               province = sheet.cell(rownum-provCounter, 2).value

               nums = []
               
               #Removes white spaces from cells with numerical values, replaces black cells with 0
               for colnum in [3, 4, 5]:
                    if len(str(sheet.cell(rownum, colnum).value).strip()) == 0:
                         nums.append(0)
                    else:
                         nums.append(int(float(str(sheet.cell(rownum, colnum).value).strip())))
               writer.writerow( (str(sheet.cell(0,1).value.encode('ascii', 'ignore')).upper(), calamityname.upper(), region.upper(), province.upper(), sheet.cell(rownum,2).value.encode('ascii', 'ignore').upper(), nums[0], nums[1], nums[2] ))
               regCounter = 1
               provCounter = 1
f.close()
print 'tapos'

#This rewrites the csv to fit the Data template required by the web2py system
with open('DAMAGED_HOUSES_web2py.csv', 'wb') as f:
     fshp = open('DAMAGED_HOUSES_shp.csv', 'rb')

     organization = 'JICA / OCD'

     #Formats date and time from the report
     reportDate = time.strptime(str(sheet.cell(5, 1).value), "%d %B %Y, %H:%M %p")
     aDate = time.strftime("%d/%m/%y", reportDate)
     aTime = time.strftime("%I:%M %p", reportDate)
     
     writer = csv.writer(f, quoting=csv.QUOTE_ALL)
     writer.writerow( ('Template', 'Series', 'Organisation', 'STD-WHO', 'STD-L0', 'STD-L1', 'STD-L2', 'STD-L3', 'STD-Lon', 'STD-Lat', 'STD-DATE', 'STD-TIME', 'IM1', 'IM2', 'IM3', 'IM4') )
          
     counter = 0
     for rownum in fshp.read().replace("\"","").split('\n'):
          row = rownum.split(',')
          if len(row) > 1 and counter > 0:
               print unicode(row[4])
               writer.writerow( ('Incidents Monitored', sheet.cell(2,1).value, organization, calamityname.upper(),'Philippines', ('REGION '+ row[2]), row[3], row[4], 'STD-Lon', 'STD-Lat', aDate, aTime, row[5], row[6], row[7].strip() )) 
          counter = 1
f.close()
print 'web2py'
