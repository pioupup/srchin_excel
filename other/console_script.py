#!/usr/bin/env python3
import sys,os,xlrd

s_path = "." # search files in path
search_text = "кардан"

if len(sys.argv) > 1:
	if len(sys.argv) > 1:	# first argument search text
		search_text = sys.argv[1]
		if len(sys.argv) > 2:
			if os.path.exists(sys.argv[2]):	# second argument search path
				s_path = sys.argv[2]
			else:
				print("Error, invalid path!")
				quit()

def search_in_path(s_path):
	list_files = [] # list for finding files
	for rootdir, dirs, files in os.walk(s_path):
		for file in files:       
			if file.split('.')[-1] == 'xls' or file.split('.')[-1] == 'xlsx':
				list_files.append(os.path.join(rootdir, file))	# append file path if it xls or xlsx
	return list_files

def srch_txt_in_xlsfile(xlsfile, search_text):
	wrkbook = xlrd.open_workbook(xlsfile)
	for sheet_nmb in range(wrkbook.nsheets):	# watching sheets in file
		sheet = wrkbook.sheet_by_index(sheet_nmb)
		for rownum in range(sheet.nrows):	# going to rows
			for colnum in range(sheet.ncols):	# going to cols
				if search_text in str(sheet.cell(rownum, colnum).value): # coincidence with searching string
					print("-" * 60)
					#print(sheet.cell(rownum, colnum).value, "[ Строка:", colnum + 1, "Столбец:", rownum + 1, "Лист:", sheet_nmb, "Файл:", xlsfile, "]")
					for cols in range(sheet.ncols):
						print(str(sheet.cell(rownum,cols).value),end=" ")
					print("\n[Column:" + str(colnum+1) + " Row:" + str(rownum+1) + " Sheet:" + str(sheet_nmb) + " Path:" + str(xlsfile) + "]")
					
for wrk_file in search_in_path(s_path):
	srch_txt_in_xlsfile(wrk_file, search_text)
