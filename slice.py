#/usr/bin/env python

def fasta_slice(name, sequence, slice=100):
	result = []
	for i, x in enumerate(sequence):
		if i > (len(sequence) - slice):
			break
		result.append(sequence[i:i+slice])
	return result

if __name__ == "__main__":
	import sys
	from Bio import SeqIO
	with open(sys.argv[1]) as infile:
		for record in SeqIO.parse(infile, 'fasta'):
			print fasta_slice(record.id, record.seq)
