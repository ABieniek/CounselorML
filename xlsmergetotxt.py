import sys
import os
from xlrd import open_workbook
from xlutils.copy import copy
import xlwt

# our files
import gpamath
import xlsmod

def printusage():
    print "usage: python ./xlsmerge inputdirectory outputname.txt outputdirectory"

def main():
    if len(sys.argv) != 4:
        printusage()
        sys.exit()
    inputdirectory = sys.argv[1]
    outputname = sys.argv[2]
    outputdirectory = sys.argv[3]

    if not os.path.isdir(inputdirectory):
        print "invalid input directory! (argument 1)"
        sys.exit()
    if not os.path.isdir(outputdirectory):
        print "invalid output directory! (argument 3)"
        sys.exit()
    if not xlsmod.isExt(outputname, [".txt"]):
        print "invalid output directory! (argument 2)"
        print "file extension should be .txt"
        sys.exit()

    outfile = open(outputdirectory + "/" + outputname, 'w')

    # make an output text file
    # write headers
    for colnum in range(len(xlsmod.headers)):
        outfile.write(xlsmod.headers[colnum])
        if (colnum != len(xlsmod.headers) - 1):
            outfile.write('\t')

    # go through files and copy all columns, except for headers on row 0
    writerow = 1    # row that we'll write to in output
    for filename in os.listdir(inputdirectory):
        if (xlsmod.isExt(filename, [".xls", ".xlsx"]) == False):
            continue
        print "copying contents of ", inputdirectory, '/', filename
        rb = open_workbook(inputdirectory + '/' + filename)
        r_sheet = rb.sheet_by_index(0)
        for rownum in range(1, r_sheet.nrows): # ignore header row
            outfile.write('\n')
            for colnum in range(0, r_sheet.ncols):
                outfile.write(str(r_sheet.cell(rownum, colnum).value))
                if (colnum != r_sheet.ncols - 1):
                    outfile.write('\t')
    # save output file
    outfile.close()

if __name__ == '__main__':
    main()
