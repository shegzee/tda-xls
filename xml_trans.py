# Use like this:
# C:\Users\Taiwo\Documents\Olusegun\work\xslx\code>workon xsltrans
# (xsltrans) C:\Users\Taiwo\Documents\Olusegun\work\xslx\code>python xml_trans.py help
# usage: xml_trans.py [-h] [-source_path SOURCE_PATH]
#                     [-destination_file DESTINATION_FILE]
#                     [-name_column NAME_COLUMN] [-value_column VALUE_COLUMN]
# xml_trans.py: error: unrecognized arguments: help

# (xsltrans) C:\Users\Taiwo\Documents\Olusegun\work\xslx\code>python xml_trans.py -h
# usage: xml_trans.py [-h] [-source_path SOURCE_PATH]
#                     [-destination_file DESTINATION_FILE]
#                     [-name_column NAME_COLUMN] [-value_column VALUE_COLUMN]

# Compile excel values

# optional arguments:
#   -h, --help            show this help message and exit
#   -source_path SOURCE_PATH
#                         Path to folder containing raw xml-formatted xsl files
#   -destination_file DESTINATION_FILE
#                         File to save results to
#   -name_column NAME_COLUMN
#                         Column of nutrient names
#   -value_column VALUE_COLUMN
#                         Column of nutrient values

# (xsltrans) C:\Users\Taiwo\Documents\Olusegun\work\xslx\code>python xml_trans.py -source_path "..\ui excel" -destination_file "uiresults.xls"

# workon xsltrans
# py xml_trans.py -source_path "..\\Fortified xls" -destination_file "RHM_Fortified.xls"

import os
from pathlib import Path
import re

# from xlutils.copy import copy
# import xlrd
from xlxmlrd import xlxmlrd as xlrd
import xlwt

def do(source_path, destination_file, name_col=1, val_col=2):
	row_offset = 7
	# name_col = 1
	# val_col = 2
	nutr_count = 27

	book = xlwt.Workbook(encoding="utf-8")
	sheet = book.add_sheet('Sheet 1')

	# files = os.listdir(path)
	i = 0
	p = Path(source_path)
	# print([x for x in p.iterdir() if x.is_dir()])
	# with os.scandir(path) as it:
	it = list(p.glob('**/*.xls'))
	print(it)

	# WRITE COLUMN HEADERS
	rb = xlrd.open_workbook(it[0])
		# wb = copy(rb)
	sheet.write(0, 0, "File name")
	this_sheet = rb.sheet_by_index(0)
	for j in range(0, nutr_count):
		nutr_value = this_sheet.cell(row_offset+j, name_col).value
		# print(nutr_value)
		sheet.write(i, j+1, nutr_value)

	# WRITE VALUES
	for each_file in it:
		print(each_file)
		# move to next row
		if "-" in each_file.name:
			print("Skipping duplicate:", each_file.name)
			continue

		i += 1
		# if each_file.is_file() and each_file.name.endswith('xsl'):
		rb = xlrd.open_workbook(each_file)
		# wb = copy(rb)
		this_sheet = rb.sheet_by_index(0)
		
		sheet.write(i, 0, each_file.name)

		# iterate over rows in source file
		for j in range(0, nutr_count):
			nutr_value = this_sheet.cell(row_offset+j, val_col).value
			if nutr_value is None or nutr_value.strip() == "":
				p = re.compile(r'\d+.*\d* *\w+')
				label_value = this_sheet.cell(row_offset+j, val_col-1).value
				if len(label_value) == 0 or label_value[0].isalpha():
					m = p.search(label_value)
					if m:
						nutr_value = m.group()
					# print(this_sheet.cell(row_offset+j, val_col-1).value)

			# print(nutr_value)
			sheet.write(i, j+1, nutr_value)

	book.save(destination_file)


if __name__ == '__main__':
	path = os.path.join('..', 'ITTtest')
	import argparse
	parser = argparse.ArgumentParser(description='Compile excel values')
	parser.add_argument('-source_path', required=False, default=path, help="Path to folder containing raw xml-formatted xsl files")
	parser.add_argument('-destination_file', required=False, default="result.xls", help="File to save results to")
	parser.add_argument('-name_column', required=False, default=1, type=int, help="Column of nutrient names")
	parser.add_argument('-value_column', required=False, default=2, type=int, help="Column of nutrient values")
	args = parser.parse_args()
	do(args.source_path, args.destination_file, args.name_column, args.value_column)

	# import sys
	# args = sys.argv
	# do(args[1], args[2])
	# path = "..\\keep\\"
	
	# main()
