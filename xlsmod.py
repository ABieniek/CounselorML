import sys
import os
from xlrd import open_workbook
from xlutils.copy import copy
from xlwt import easyxf

# our files
import gpamath

def printusage():
    print "usage: python ./wow directory"

# return true if the given filename has one of the extensions in the list
def isExt(filename, extList):
    for ext in extList:
        if (filename.find(ext) == -1):
            continue
        if (filename.find(ext) + len(ext) == len(filename)):
            return True
    return False

# check arguments
if len(sys.argv) > 2:
    printusage()
    sys.exit()
directory = "."
if len(sys.argv) != 1:
    directory = sys.argv[1]

def clearsheet(writesheet, nrows, ncols):
    for i in xrange(nrows, -1, -1):
        for j in xrange(ncols, -1, -1):
            writesheet.write(i, j, "")
    print "done clearing"

def pushCols(readsheet, writesheet, startcol, dist):
    # copy current columns up to the column before startcol
    if startcol >= readsheet.nrows:
        print "starting column is not in range!"
        return None
    for colnum in range(startcol):
        for rownum in range(readsheet.nrows):
            writesheet.write(rownum, colnum, readsheet.cell(rownum, colnum).value)

for filename in os.listdir(directory):
    if (isExt(filename, [".xls", ".xlsx"]) == False):
        continue
    rb = open_workbook(directory + '/' + filename)
    r_sheet = rb.sheet_by_index(0)
    wb = copy(rb)
    w_sheet = wb.get_sheet(0)
    clearsheet(w_sheet, r_sheet.nrows, r_sheet.ncols)
    pushCols(r_sheet, w_sheet, 3, 0)
    #cb = copy(wb)
    wb.save(directory + '/temp' + filename)
