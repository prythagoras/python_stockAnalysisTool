'''
Date: March 22 2020
Author: Prithivi Maruthachalam
Title: Personalized Command Line Stock Analysis Tool
'''

import numpy as np
import pandas as pd
import sys

def interpolate(start,end,num):
	percentage = float(((float(num)-float(start))/(float(end)-float(start))))
	return(percentage)

def extrapolate(start,end,start2,end2,num):
	percentage = interpolate(start,end,num)
	value = (float(end2)-float(start2)+1.0)*percentage
	return float(value)
	
	
toBeNormalized = {"P/D":(100,10),
"P/E":(30,1),
"P/B":(2.5,0.25),
"(P/E) * (P/B)":(40,10),
"Current Ratio":(1.5,0.1),
"Debt equity ratio":(1,0.1)}

#function for normalisation
def normalize(value,param):
	POP = toBeNormalized[param] 
	if(value >= POP[0]):
		return 1
	if(value <= POP[1]):
		return 10
	return extrapolate(POP[0],POP[1],1,10,value)

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
		break
	if(entered >= 0 and (entered in possibles)):
		weight = float(input(">> Enter weight for " + headers[entered] + " :  "))
		weightsAndParams[str(headers[entered])] = weight
	if(entered not in possibles and entered != -1):
		print("[WARNING]: Entered value not in list.")

print("**************************************************")
print()
noOfStocks = data.shape[0]
print("Analysing " + str(noOfStocks) + " stocks from " + oldFileName)

for i in range(noOfStocks):
	score = 0
	for key,val in weightsAndParams.items():
		if(key in toBeNormalized.keys()):
			normalizedValue = normalize(data[key][i],key)
			data[str(key) + ".1"][i] = normalizedValue
			score += normalizedValue * val
		else:
			score += data[key][i] * val
	data["Score"][i] = score
		
print("[INFO]: Analysis complete!")

data.sort_values(by=["Score"],inplace=True,ascending=False)
data.to_excel(newFileName)
print("[INFO]: Analysis written to " + newFileName)
