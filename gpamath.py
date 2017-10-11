from xlrd import open_workbook
from xlutils.copy import copy
from xlwt import easyxf

# dictionary that pairs a letter grade with a value
gradeVals = {'A+': 4.00, 'A': 4.00, 'A-': 3.67, 'B+': 3.33, 'B': 3.00, \
            'B-': 2.67, 'C+': 2.33, 'C': 2.00, 'C-': 1.67, 'D+': 1.33, \
            'D': 1.00, 'D-': 0.67, 'F': 0.00}

def calcRowAverage(sheet, rownum):
    count = 0
    sumgrade = 0
    for j in range(sheet.ncols):
        if (sheet.cell(0, j).value in gradeVals):
            count += sheet.cell(rownum, j).value    # increase number of students counted
            sumgrade +=  sheet.cell(rownum, j).value \
            * gradeVals[sheet.cell(0, j).value]     # increase total score, weighted by grade value
    return sumgrade / count
