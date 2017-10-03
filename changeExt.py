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
lenoldext = len(oldext)
newext = sys.argv[2]
lennewext = len(newext)
directory = sys.argv[3]

for filename in os.listdir(directory):
    if (filename.find(oldext) + lenoldext == len(filename)):
        extidx =  filename.find(oldext)
        os.rename(directory + '/' + filename, directory + '/' + filename[0:extidx] + newext)
