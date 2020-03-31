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

import numpy as np
import pandas as pd
import sys


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
	print(POP,newPOP)
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
	
data = pd.read_excel(oldFileName)
headers = data.keys()


weightsAndParams = dict()
possibles = list()
print("**************************************************")
print("Enter choices for parameter and associated weight")
for r,parameter in enumerate(headers):
	if not data[parameter].isnull().values.any():
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
			weightsAndParams["(P/E) * (P/B)"] = int(input(">> Enter weight for (P/E) * (P/B): "))
			LOOPING = False
		except ValueError:
			print("[WARNING]: Please enter a valid value!")
			print()
			LOOPING = True
	
print("**************************************************")
print()
noOfStocks = data.shape[0]
print("Analysing " + str(noOfStocks) + " stocks from " + oldFileName)
	
print(weightsAndParams)
	
for i in range(noOfStocks):
	score = 0
	for key,val in weightsAndParams.items():
		if(key in toBeNormalized.keys()):
			normalizedValue = normalize(data[key][i],key)
			try:
				data[str(key) + ".1"][i] = normalizedValue
			except KeyError:
				data[str(key)][i] = normalizedValue
			score += normalizedValue * val
		else:
			score += data[key][i] * val
		print(key,val,normalizedValue)
	data["Score"][i] = score

	
		
print("[INFO]: Analysis complete!")

data.sort_values(by=["Score"],inplace=True,ascending=False)
data.to_excel(newFileName)
print("[INFO]: Analysis written to " + newFileName)
