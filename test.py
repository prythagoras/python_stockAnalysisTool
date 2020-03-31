import xlrd
loc = ("stockAnalysis.xlsx") 
  
# To open Workbook 
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
 

headers = list()
for i in range(sheet.ncols):
	headers.append(sheet.cell_value(0,i))
	


data = list()
for col in range(sheet.ncols):
	rows = list()
	for row in range(sheet.nrows):
		rows.append(sheet.cell_value(row,col))
	data.append(rows)

print(data[3cle][1])
