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

# also includes term header before the rest in one year, but I'll omit that
headers = ["Term", "Subject","Course", "CRN", "Course Title", "Total Grades", \
    "A+", "A", "A-", "B+", "B", "B-", "C+", "C", "C-",\
    "D+", "D", "D-", "F", "W", "Average Grade", "Primary Instructor"]

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
    for colnum in range(startcol, readsheet.ncols):
        for rownum in range(readsheet.nrows):
            writesheet.write(rownum, colnum+dist, readsheet.cell(rownum, colnum).value)

def spacecols(readsheet, writesheet):
    readcol = 0
    for headnum in range(len(headers)):
        headertitle = readsheet.cell(0, readcol).value.encode('ascii','ignore')
        if headers[headnum] in headertitle:
            for rownum in range(readsheet.nrows):
                writesheet.write(rownum, headnum, readsheet.cell(rownum, readcol).value)
            readcol += 1
        else:
            writesheet.write(0, headnum, headers[headnum])


for filename in os.listdir(directory):
    if (isExt(filename, [".xls", ".xlsx"]) == False):
        continue
    rb = open_workbook(directory + '/' + filename)
    r_sheet = rb.sheet_by_index(0)
    wb = copy(rb)
    w_sheet = wb.get_sheet(0)
    clearsheet(w_sheet, r_sheet.nrows, r_sheet.ncols)
    spacecols(r_sheet, w_sheet)
    #cb = copy(wb)
    wb.save(directory + '/temp' + filename)
