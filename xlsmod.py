import sys
import os
from xlrd import open_workbook
from xlutils.copy import copy
from xlwt import easyxf

# our files
import gpamath

def printusage():
    print "usage: python ./xlsmod directory"

# return true if the given filename has one of the extensions in the list
def isExt(filename, extList):
    for ext in extList:
        if (filename.find(ext) == -1):
            continue
        if (filename.find(ext) + len(ext) == len(filename)):
            return True
    return False

# also includes term header before the rest in one year, but I'll omit that
headers = ["Term", "Subject","Course", "CRN", "Course Title", "Total Grades", \
    "A+", "A", "A-", "B+", "B", "B-", "C+", "C", "C-",\
    "D+", "D", "D-", "F", "W", "Average Grade", "Primary Instructor"]

# http://pythoncentral.io/how-to-check-if-a-string-is-a-number-in-python-including-unicode/
def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        pass
    try:
        import unicodedata
        unicodedata.numeric(s)
        return True
    except (TypeError, ValueError):
        pass

# searches through str and tries to find a possible year number of length n
def guessyear(str, n):
    strtemp = ""
    strings = []
    consec = 0
    for c in str:
        if (is_number(c)):
            strtemp += c
            consec+=1
        else:
            if (consec == n):
                strings.append(strtemp)
            strtemp = ""
            consec = 0
    if (consec == n):
        strings.append(strtemp)
    for s in strings:
        if int(s) in range(1900, 3000):
            return int(s)
    return 0

def guessseason(filename):
    if "FALL" in filename.upper(): return "FALL"
    elif "WINTER" in filename.upper(): return "WINTER"
    elif "SPRING" in filename.upper(): return "SPRING"
    elif "SUMMER" in filename.upper(): return "SUMMER"
    return ""

def guessterm(filename):
    season = guessseason(filename)
    if (season == ""): print "no season found in filename!"
    year = guessyear(filename, 4)
    if (year == 0): print "no year found in filename!"
    return season, year

# write "" to every populated cell in a sheet
def clearsheet(writesheet, nrows, ncols):
    for i in xrange(nrows, -1, -1):
        for j in xrange(ncols, -1, -1):
            writesheet.write(i, j, "")

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

def fillterms(readsheet, writesheet, filename):
    for rownum in range(1, readsheet.nrows):
        for colnum in range(readsheet.ncols):
            if readsheet.cell(0, colnum).value == "Term":
                season, year = guessterm(filename)
                writesheet.write(rownum, colnum, season + " " + str(year))


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

def main():
    # check arguments
    if len(sys.argv) != 2:
        printusage()
        sys.exit()
    directory = sys.argv[1]
    if not os.path.isdir(directory):
        printusage()
        sys.exit()

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
        # first fill, fill empty letter grade columns and terms
        rb = open_workbook(directory + '/' + filename)
        r_sheet = rb.sheet_by_index(0)
        wb = copy(rb)
        w_sheet = wb.get_sheet(0)
        fillgrades(r_sheet, w_sheet)
        fillterms(r_sheet, w_sheet, filename)
        wb.save(directory + '/' + filename)
        # second fill, now I have all values to calculate GPA with
        rb = open_workbook(directory + '/' + filename)
        r_sheet = rb.sheet_by_index(0)
        wb = copy(rb)
        w_sheet = wb.get_sheet(0)
        fillaverages(r_sheet, w_sheet)
        wb.save(directory + '/' + filename)

if __name__ == '__main__':
    main()
