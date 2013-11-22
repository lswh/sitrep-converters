from xlutils.copy import copy # http://pypi.python.org/pypi/xlutils
from xlrd import open_workbook # http://pypi.python.org/pypi/xlrd
from xlwt import easyxf # http://pypi.python.org/pypi/xlwt
import sys
import csv

rb = open_workbook('B_Casualties2_test.xls',formatting_info=True)
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

wb.save('B_Casualties2_test.xls')

rb = open_workbook('B_Casualties2_test.xls',formatting_info=True)
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
wb.save('B_Casualties2_test.xls')
