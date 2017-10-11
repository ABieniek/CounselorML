import sys
import os
from xlrd import open_workbook
from xlutils.copy import copy
from xlwt import easyxf

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

# dictionary that pairs a letter grade with a value
gradeVals = {'A+': 4.00, 'A': 4.00, 'A-': 3.67, 'B+': 3.33, 'B': 3.00, \
            'B-': 2.67, 'C+': 2.33, 'C': 2.00, 'C-': 1.67, 'D+': 1.33, \
            'D': 1.00, 'D-': 0.67, 'F': 0.00}

# check arguments
if len(sys.argv) > 2:
    printusage()
    sys.exit()
directory = "."
if len(sys.argv) != 1:
    directory = sys.argv[1]

def calcRowAverage(sheet, rownum):
    count = 0
    sumgrade = 0

    for j in range(sheet.ncols):
        if (sheet.cell(0, j).value in gradeVals):
            count += sheet.cell(rownum, j).value
            sumgrade +=  sheet.cell(rownum, j).value \
            * gradeVals[sheet.cell(0, j).value]
    return sumgrade / count

def pushCols(readsheet, writesheet, startidx, dist):
    ncols = readsheet.ncols
    # write new values to the columns some distance away
    for j in xrange(ncols, startidx, -1):
        for i in range(readsheet.nrows):
            print j + dist
            writesheet.write(i, j+dist, readsheet.cell(i, j))
    # clear the columns that got copied and not overwritten
    for j in range(startidx, distance):
        for i in range(readsheet.nrows):
            writesheet.write(i, j, "")

for filename in os.listdir(directory):
    if (isExt(filename, [".xls", ".xlsx"]) == False):
        continue
    rb = open_workbook(directory + '/' + filename)
    r_sheet = rb.sheet_by_index(0)
    wb = copy(rb)
    w_sheet = wb.get_sheet(0)
    pushCols(r_sheet, w_sheet, 5, 1)
    cb = copy(wb)
    cb.save(directory + '/temp' + filename)
