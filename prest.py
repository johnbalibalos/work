#!/usr/bin/env python
"""
Takes a fasta file (taken from uniprot, etc) and does a 
blastp search on a sliding window of aa sequence. Throws
out sequence identities over 60%.

Usage: python antibody.py <fasta>
"""

import numpy
from Bio import SeqIO
from Bio.Blast import NCBIWWW
from Bio.Blast import NCBIXML
import csv
#import matplotlib
import sys
from collections import deque
import slice

if len(sys.argv)==1:
    sys.exit('No file specified. Exiting.')
infile = sys.argv[1]

file = open(infile, 'rU')
for record in SeqIO.parse(infile, 'fasta'):
	print record.id
	outfile = record.id + '.csv'
    #print record.seq
file.close()

def blastp(fasta_string):
	SLICE = 100
	E_VALUE_THRESH = 0.004
	SCORE = 0
	MATCHES = 10
	IDENTITY = 60
	blast_handle = NCBIWWW.qblast('blastp', 'nr', fasta_string, alignments=1, entrez_query='txid9606 [ORGN]')
	blast_parse = NCBIXML.parse(blast_handle)
	record = blast_parse.next()
	for i,alignment in enumerate(record.alignments):
		for hsp in alignment.hsps:
			if hsp.identities < IDENTITY:
				#print "***Alignment***"
				#print "sequence:", alignment.title
				#print "length:", alignment.length
				#print "E-value:", hsp.expect
				#print "Score:", hsp.score
				#print hsp.query[:50] + "..."
				#print hsp.match[:50] + "..."
				#print hsp.sbjct[:50] + "..."
				yield [fasta_string, alignment.title, alignment.length, hsp.expect, hsp.score, hsp.identities, hsp.positives]
#blastp(record.seq)
#print blastp(record.seq)

with open(outfile, 'wb') as csvfile:
	Data = []
	for i,x in enumerate(fasta_slice(record.id, record.seq)):
		data = []
		for y in blastp(x):
			data.append(y)
		writer = csv.writer(csvfile, delimiter='\t')
		writer.writerows(data)
