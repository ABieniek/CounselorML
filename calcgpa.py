import sys
import os
from xlrd import open_workbook

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
    return True

for filename in os.listdir(directory):
    if (isExt(filename, [".xls", ".xlsx"]) == False):
        continue
    wb = open_workbook(directory + '/' + filename)
    sheet = wb.sheet_by_index(0)
    for i in range(1, sheet.nrows):
        sumgrade = 0
        count = 0
        for j in range(sheet.ncols):
            # if the column header we're looking at has a letter grade \
            # this column contains numbers of students with that score
            if (sheet.cell(0, j).value in gradeVals):
                count += sheet.cell(i, j).value
                sumgrade +=  sheet.cell(i, j).value \
                * gradeVals[sheet.cell(0, j).value]
        print (sumgrade / count)
