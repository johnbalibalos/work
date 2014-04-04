#!usr/bin/env python
import csv
import re
import os
import sys
import collections
from bioservices import UniProt

if len(sys.argv)==1:
    print 'No file specified. Exiting.'
    sys.exit()

if not os.path.exists(sys.argv[1]):
    sys.stderr.write('Error: File %s was not found! ' % sys.argv[1])
    sys.exit()

if not sys.argv[1].endswith('.csv'):
    sys.stderr.write('File is not a csv file.')
    sys.exit()

#opens annotation file
with open('082812annotation.txt', 'rU') as infile:
    reader = csv.reader(infile, delimiter='\t')
    mydict = {}
    for rows in reader:
        k = rows[0]
        v = rows[:]
        mydict[k] = v
    #print mydict

#regex the (+1) after protein name
def regex(name):
    pattern = r'_HUMAN.*'
    replacement = r'_HUMAN'
    outname = re.sub(pattern, replacement, name)
    return outname

#opens protein data file
with open(sys.argv[1], 'rU') as data:
    datafile = csv.reader(data, delimiter=',')
    header = datafile.next()
    data = []
    for entries in datafile:
        data.append(entries)
    SAMPLE_NAMES = header[1:]
    datadict = {}
    for idx,sample in enumerate(SAMPLE_NAMES):
        proteindata = {}
        for x in data:
            k = regex(x[0])
            v = x[idx+1]
            proteindata[k] = v
        datadict[sample] = proteindata
    #print proteindata.keys()
    #print mydata.keys()
