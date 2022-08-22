#!/usr/bin/env python3
import os
import pandas as pd
import openpyxl
import glob
import xlsxwriter
import subprocess


## Get filepath from user, get battery sample name from filepath

filepath = input('\nDrop your .xlsx DCIR result file here:\n\n')
filepath = filepath.strip()
lastSlash = filepath.rindex("/")


## Scan folder for other other result files
sampleFolder = filepath[:lastSlash]
xlsx_files = glob.glob(os.path.join(sampleFolder, "*.xlsx"))


## Loop over the list of result files

output = []

for f in xlsx_files:

	## Read in result file
	df = pd.read_excel(f, engine='openpyxl', header=1)

	## Identify correct columns for ES, thermocouple, and GG temp
	columns = []
	columns = df.columns.get_indexer(['ES', 'Temp 2', 'Temperature (0x06)'])

	## Conditional statement, to handle different Maccor thermocouple
	## labeling
	if -1 in columns:
		print(columns)
		columns = df.columns.get_indexer(['ES', 'Temp 1', 'Temperature (0x06)'])

	EScolumn = columns[0]
	tempTherm = columns[1]
	tempGG = columns[2]

	## Read in result file again, this time knowing which columns to use
	df = pd.read_excel(f, engine='openpyxl', header=1, usecols=columns) 

	sampleNameIndex = f.rindex("/")
	sampleName = f[sampleNameIndex+1:-9]
	print(sampleName)


	## Filter ES column down to only '129' indicators
	step129 = df['ES'] == 129
	dataFrame = df[step129]
	print(columns)
	print(dataFrame)
	##dataFrame = dataFrame.astype(float, errors='raise')

	neg20Therm = dataFrame.iloc[2,1]
	neg20GG = dataFrame.iloc[2,2]

	if neg20GG > -17:
		

	neg10Therm = dataFrame.iloc[3,1]
	neg10GG = dataFrame.iloc[3,2]

	zeroTherm = dataFrame.iloc[4,1]
	zeroGG = dataFrame.iloc[4,2]

	pos25Therm = dataFrame.iloc[5,1]
	pos25GG = dataFrame.iloc[5,2]

	pos45Therm = dataFrame.iloc[6,1]
	pos45GG = dataFrame.iloc[6,2]

	pos60Therm = dataFrame.iloc[7,1]
	pos60GG = dataFrame.iloc[7,2]
		

	neg20TempAcc = neg20GG - neg20Therm
	neg10TempAcc = neg10GG - neg10Therm
	zeroTempAcc = zeroGG - zeroTherm
	pos25TempAcc = pos25GG - pos25Therm
	pos45TempAcc = pos45GG - pos45Therm
	pos60TempAcc = pos60GG - pos60Therm

	output.append([sampleName, neg20TempAcc, neg20Therm, neg20GG, neg10TempAcc,\
	neg10Therm, neg10GG, zeroTempAcc, zeroTherm, zeroGG, pos25TempAcc,\
	pos25Therm, pos25GG, pos45TempAcc, pos45Therm, pos45GG, pos60TempAcc,\
	pos60Therm, pos60GG])

	print('\n')

outputDF = pd.DataFrame(output, columns = ['Sample', 'TempAcc (-20°C)', 'Thermo',\
'GG', 'TempAcc (-10°C)', 'Thermo', 'GG', 'TempAcc (0°C)', 'Thermo', 'GG',\
'TempAcc (25°C)', 'Thermo', 'GG', 'TempAcc (45°C)', 'Thermo', 'GG', 'TempAcc (60°C)',\
'Thermo', 'GG'])

##outputFileName = sampleName[:-2]
##outputFileName = outputFileName + '.xlsx'
outputDF.to_excel('output.xlsx', engine = 'xlsxwriter')

subprocess.run(['open', 'output.xlsx'], check=True)
