import sys
import os

def printusage():
    print "usage: python ./changeExt .oldext .newext directory"

if len(sys.argv) != 4:
    printusage()
    sys.exit()

if (sys.argv[1][0] != '.' or sys.argv[2][0] != '.'):
    printusage()

oldext = sys.argv[1]
newext = sys.argv[2]
directory = sys.argv[3]

for filename in os.listdir(directory):
    if (filename.find(newext) == -1 and filename.find(oldext) != -1):
        extidx =  filename.find(oldext)
        os.rename(filename, filename[0:extidx] + newext)
