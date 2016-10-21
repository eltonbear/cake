import pyexcel as pe
import os

def readSheet(filePath):
	
	fileBaseName = os.path.basename(filePath)
	fileName, fileExtension = os.path.splitext(fileBaseName) ############################################# FIRST SHEET NAME????????????/
	if fileExtension == '.csv':
		return {'error': '.csv file type does not support multiple sheets. Please save the file in .xlsx.'}
	try:
		# Get workbook that has two sheets
		book = pe.get_book(file_name=filePath) 
		
		if 'Alignment' not in book.sheet_names():
			return {'error': 'File: ' + filePath + ' missing "Alignment" worksheet.'}
		if fileName not in book.sheet_names():
			return {'error': 'File: ' + filePath + ' missing "' + fileName + '" worksheet.'}

	except NotImplementedError as error:
		return {'error': 'No source found for ' + filePath}
	except FileNotFoundError as error:
		return {'error': error.args[1] + ': ' + filePath}

	partsTitleRow = 1
	partSheet = book[fileName]
	alignmentSheet = book['Alignment']
	partSheet.name_columns_by_row(partsTitleRow)
	alignmentSheet.name_columns_by_row(0)

	partsTitleRow = -1
	productInfoC = 0
	productInfoR = 0
	emptyRow = 1
	requiredPartTitles = ['Ref Des', 'Part Number', 'X-location', 'Y-location', 'Rotation', 'Layer']
	requriedAlignTitles = ['Layer',	'x1', 'y1', 'x2', 'y2']
	partSheet.colnames = list(map(str.strip, partSheet.colnames))
	alignmentSheet.colnames = list(map(str.strip, alignmentSheet.colnames))
	partTitles = partSheet.colnames
	alignmentTitles = alignmentSheet.colnames
	emptyRowList = partSheet.row[emptyRow]
	missingPartColumns = []
	missingAlignColumns = []


	# Check if required columns all exist for partSheet and alignment sheet
	for requireTitle in requiredPartTitles:
		if requireTitle not in partTitles:
			missingPartColumns.append(requireTitle)
	if missingPartColumns:	
		return{'error': 'File missing column of "' + ','.join(missingPartColumns) + '" in parts sheet!'}
	for requireTitle in requriedAlignTitles:
		if requireTitle not in alignmentTitles:
			missingAlignColumns.append(requireTitle)
	if missingAlignColumns:	
		return{'error': 'File missing column of "' + ','.join(missingAlignColumns) + '" in alignment sheet!'}

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
		partsDictionaries = partSheet.to_records()
		alignmentDictsTemp = alignmentSheet.to_records()
	except:
		return {'error': 'File format incorrect: ' + filePath}

	alignmentDictionaries = {}
	for alignDict in alignmentDictsTemp:
		alignmentDictionaries[alignDict['Layer']] = {'p1': (alignDict['x1'], alignDict['y1']), 'p2': (alignDict['x2'], alignDict['y2'])}

	return {'productName': productName, 'data': partsDictionaries, 'fiducials': alignmentDictionaries}





