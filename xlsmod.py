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

# write "" to every populated cell in a sheet
def clearsheet(writesheet, nrows, ncols):
    for i in xrange(nrows, -1, -1):
        for j in xrange(ncols, -1, -1):
            writesheet.write(i, j, "")

def pushcols(readsheet, writesheet, startcol, dist):
    # copy current columns up to the column before startcol
    if startcol >= readsheet.nrows:
        print "starting column is not in range!"
        return None
    for colnum in range(startcol):
        for rownum in range(readsheet.nrows):
            writesheet.write(rownum, colnum, readsheet.cell(rownum, colnum).value)
    # once we hit startcol, write all of the columns in startcol some distance ahead
    for colnum in range(startcol, readsheet.ncols):
        for rownum in range(readsheet.nrows):
            writesheet.write(rownum, colnum+dist, readsheet.cell(rownum, colnum).value)

def spacecols(readsheet, writesheet):
    readcol = 0
    for headnum in range(len(headers)):
        headertitle = readsheet.cell(0, readcol).value.encode('ascii','ignore')
        if headers[headnum] in headertitle:
            writesheet.write(0, headnum, headers[headnum])
            for rownum in range(1, readsheet.nrows):
                writesheet.write(rownum, headnum, readsheet.cell(rownum, readcol).value)
            readcol += 1
        else:
            writesheet.write(0, headnum, headers[headnum])

def stripbadrows(readsheet, writesheet):
    writerow = 0
    for rownum in range(readsheet.nrows):
        # check if our row is garbage
        skip = False
        for colnum in range(readsheet.ncols):
            if readsheet.cell(rownum, colnum).value == "N/A":
                skip = True
                break
        if skip == False:
            for colnum in range(readsheet.ncols):
                writesheet.write(writerow, colnum, readsheet.cell(rownum, colnum).value)
            writerow+=1

def fillgrades(readsheet, writesheet):
    for rownum in range(1, readsheet.nrows):
        for colnum in range(readsheet.ncols):
            if readsheet.cell(0, colnum).value in headers[headers.index("A+"):headers.index("Average Grade")]:
                if readsheet.cell(rownum, colnum).value == "":
                    writesheet.write(rownum, colnum, 0)

def fillaverages(readsheet, writesheet):
    for rownum in range(1, readsheet.nrows):
        for colnum in range(readsheet.ncols):
            if readsheet.cell(0, colnum).value == "Total Grades":
                writesheet.write(rownum, colnum, gpamath.calcrowcount(readsheet, rownum))
            elif readsheet.cell(0, colnum).value == "Average Grade":
                writesheet.write(rownum, colnum, gpamath.calcrowaverage(readsheet, rownum))

for filename in os.listdir(directory):
    if (isExt(filename, [".xls", ".xlsx"]) == False):
        continue
    print "rewriting ", directory, '/', filename
    #  make a temp file that has the desired format
    rb = open_workbook(directory + '/' + filename)
    r_sheet = rb.sheet_by_index(0)
    wb = copy(rb)
    w_sheet = wb.get_sheet(0)
    clearsheet(w_sheet, r_sheet.nrows, r_sheet.ncols)
    spacecols(r_sheet, w_sheet)
    wb.save(directory + '/' + filename)
    # strip bad rows
    rb = open_workbook(directory + '/' + filename)
    r_sheet = rb.sheet_by_index(0)
    wb = copy(rb)
    w_sheet = wb.get_sheet(0)
    stripbadrows(r_sheet, w_sheet)
    wb.save(directory + '/' + filename)
    # first fill, fill empty letter grade columns
    rb = open_workbook(directory + '/' + filename)
    r_sheet = rb.sheet_by_index(0)
    wb = copy(rb)
    w_sheet = wb.get_sheet(0)
    fillgrades(r_sheet, w_sheet)
    wb.save(directory + '/' + filename)
    # second fill, now I have all values to calculate GPA with
    rb = open_workbook(directory + '/' + filename)
    r_sheet = rb.sheet_by_index(0)
    wb = copy(rb)
    w_sheet = wb.get_sheet(0)
    fillaverages(r_sheet, w_sheet)
    wb.save(directory + '/' + filename)
