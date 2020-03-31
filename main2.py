'''
Date: March 22 2020
Author: Prithivi Maruthachalam
Title: Personalized Command Line Stock Analysis Tool
'''

'''
Using the toBeNormalized table
Enter each coloumn to be normalized in the following format, one on each line
"<Name of scrip>:( (lowest given value , highest given value) , (start value to be converted , end value to be converted) )"
'''	
toBeNormalized = {"P/D":((100,10),(1,10)),
"Dividend yield / years":((0,10),(0,10)),
"P/E":((30,1),(1,10)),
"P/B":((2.5,0.25),(1,10)),
"(P/E) * (P/B)":((40,10),(1,10)),
"Current Ratio":((1.5,0.1),(1,10)),
"Debt equity ratio":((1,0.1),(1,10))}

import xlrd
import xlwt


def interpolate(start,end,num):
	percentage = float(((float(num)-float(start))/(float(end)-float(start))))
	return(percentage)

def extrapolate(start,end,start2,end2,num):
	percentage = interpolate(start,end,num)
	value = ((float(end2)-float(start2))*percentage) + start2
	return float(value)
	


#function for normalisation
def normalize(value,param):
	POP,newPOP = toBeNormalized[param] 
	if(POP[0] > POP[1]):
		if(value >= POP[0]):
			return newPOP[0]
		if(value <= POP[1]):
			return newPOP[1]
	else:
		print("here")
		if(value <= POP[0]):
			return newPOP[0]
		if(value >= POP[1]):
			return newPOP[1]
	return extrapolate(POP[0],POP[1],newPOP[0],newPOP[1],value)

try:
	oldFileName = sys.argv[1]
except:
	oldFileName = 'stockAnalysis.xlsx'
	
try:
	newFileName = sys.argv[2]
except:
	newFileName = 'analysedStock.xlsx'
	
	
loc = (oldFileName)
wb = xlrd.open_workbook(loc)	
sheet = wb.sheet_by_index(0)

headers = list()
for i in range(sheet.ncols):
	headers.append(sheet.cell_value(0,i))
	
data = list()
for row in range(sheet.nrows):
	cols = list()
	for col in range(sheet.ncols):
		cols.append(sheet.cell_value(row,col))
	data.append(cols)
		
		

weightsAndParams = dict()
possibles = list()
print("**************************************************")
print("Enter choices for parameter and associated weight")
for r,parameter in enumerate(headers):
	#get index of data in headers then check if the value is null for the first coloumn
	if sheet.cell_value(1,r):
		print(str(r)," -- ",parameter)
		possibles.append(r)
print("-1"," -- ","QUIT")

entered = -2
while(entered != -1):
	print()
	try:
		entered = int(input(">> Enter choice: "))
	except ValueError:
		print("[WARNING]Please enter a number only!")
		continue 
	try:
		if(entered >= 0 and (entered in possibles)):
			weight = float(input(">> Enter weight for " + headers[entered] + " :  "))
			weightsAndParams[str(headers[entered])] = weight
		if(entered not in possibles and entered != -1):
			print("[WARNING]: Entered value not in list.")
	except ValueError:
		print("[WARNING]: Please enter a valid value!")
		continue

if "P/E" in weightsAndParams.keys() and "P/B" in weightsAndParams.keys():
	LOOPING = True
	while LOOPING:
		try:
			weightsAndParams["(P/E) * (P/B)"] = float(input(">> Enter weight for (P/E) * (P/B): "))
			LOOPING = False
		except ValueError:
			print("[WARNING]: Please enter a valid value!")
			print()
			LOOPING = True

print("**************************************************")
print()
noOfStocks = sheet.nrows - 1
print("Analysing " + str(noOfStocks) + " stocks from " + oldFileName)	

workbook = xlwt.Workbook()
sheet2 = workbook.add_sheet('test')

rowOfThis = headers.index("(P/E) * (P/B)")
for i in range(1,sheet.nrows):
	data[i][rowOfThis] = data[i][headers.index("P/E")] * data[i][headers.index("P/B")]

for i in range(noOfStocks):
	score = 0
	for key,val in weightsAndParams.items():
		keyIndex = headers.index(key)
		if(key in toBeNormalized.keys()):
			normalizedValue = normalize(sheet.cell_value(i+1,keyIndex),key)
			newKeyIndex = headers.index(str(key)+"*")
			data[i+1][newKeyIndex] = normalizedValue
			score += normalizedValue * val
		else:
			score += (sheet.cell_value(i+1,keyIndex) * val)
	data[i+1][headers.index("Score")] = score

		
print("[INFO]: Analysis complete!")
for i,row in enumerate(data):
	for r,col in enumerate(data[i]):
		sheet2.write(i,r,data[i][r])
		
print(headers.index("Score"))
data.sort(key=lambda x:x[19],reverse=True)
for row in data:
	print(row)
	print len(row)


workbook.save(newFileName)
print("[INFO]: Analysis written to " + newFileName)
