import pyexcel as pe
import os
import re

def readSheet(filePath):
	try:
		# Get workbook
		book = pe.get_book(file_name=filePath) 
		# Get sheet name in list
		sheetName = book.sheet_names()
		# error if there are more than 1 sheet in this book
		if len(sheetName) > 1:
			return {'error': 'There are more than one sheet.'}

	except NotImplementedError as error:
		return {'error': 'No source found for ' + filePath}
	except FileNotFoundError as error:
		return {'error': error.args[1] + ': ' + filePath}

	partsTitleRow = 1
	# Get sheet
	partSheet = book[sheetName[0]]
	partSheet.name_columns_by_row(partsTitleRow)

	partsTitleRow = -1
	productInfoC = 0
	productInfoR = 0
	emptyRow = 1
	requiredPartTitles = ['Ref Des', 'Part Number', 'X-location', 'Y-location', 'Rotation', 'Layer']
	partSheet.colnames = list(map(str.strip, partSheet.colnames))
	partTitles = partSheet.colnames
	emptyRowList = partSheet.row[emptyRow]
	missingPartColumns = []

	# Check if required columns all exist for partSheet and alignment sheet
	for requireTitle in requiredPartTitles:
		if requireTitle not in partTitles:
			missingPartColumns.append(requireTitle)
	if missingPartColumns:	
		return{'error': 'File missing column of "' + ','.join(missingPartColumns) + '" in parts sheet!'}

	# Check if the empty row in parts sheet is a list of strings
	if not all(isinstance(e, str) for e in emptyRowList):
		return{'error': 'File format incorrect: ' + filePath}
	# Check if there an empty at row 2
	spaceString = ''.join(emptyRowList)
	if type(spaceString) is not str or spaceString.strip() != '':
		return{'error': 'File format incorrect: ' + filePath}

	try:
		productionInfo = partSheet.row[productInfoR][productInfoC].split( )
		productName = productionInfo[0]
		del partSheet.row[emptyRow]
		del partSheet.row[productInfoR]
		partsAndFIDdictionaries = partSheet.to_records()
	except:
		return {'error': 'File format incorrect: ' + filePath}

	# Separate parts and FIDs into 2 lists
	partsDictionaries = []
	alignmentDictsTemp = []
	for part in partsAndFIDdictionaries:
		if part['Ref Des'][:3] == 'FID':
			alignmentDictsTemp.append(part)
		elif part['Ref Des'] != '':
			partsDictionaries.append(part)

	# Separate FIDs by layers --> {'L1': [D1, D2], 'L2':[D1 D2], ....}, D1 = {'Ref Des': 'FID#', .............} 
	alignmentDictionariesByLayer = {}
	for fidDict in alignmentDictsTemp:
		if fidDict['Layer'] not in  alignmentDictionariesByLayer:
			alignmentDictionariesByLayer[fidDict['Layer']] = [fidDict]
		else:
			alignmentDictionariesByLayer[fidDict['Layer']].append(fidDict)

	# Rerrange alignment fid pattern
	alignmentDictionaries = {}		
	for fidLayer in alignmentDictionariesByLayer:
		# Check the number of fid per layer
		if len(alignmentDictionariesByLayer[fidLayer]) != 2:
			return {'error': 'The number of fiducial incorrect in layer: ' + fidLayer}
		fid1 = alignmentDictionariesByLayer[fidLayer][0]
		fid2 = alignmentDictionariesByLayer[fidLayer][1]
		FIDname1 = fid1['Ref Des']
		FIDname2 = fid2['Ref Des']
		if int(re.findall(r'\d+',FIDname1)[0]) > int(re.findall(r'\d+',FIDname2)[0]):
			alignmentDictionaries[fidLayer] = {'p1': (fid2['X-location'], fid2['Y-location']), 'p2': (fid1['X-location'], fid1['Y-location'])}
		else:
			alignmentDictionaries[fidLayer] = {'p1': (fid1['X-location'], fid1['Y-location']), 'p2': (fid2['X-location'], fid2['Y-location'])}

	return {'productName': productName, 'data': partsDictionaries, 'fiducials': alignmentDictionaries}






