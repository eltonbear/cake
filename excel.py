import pyexcel as pe
import warnings

def readSheet(filePath):
	titleRow = 1
	try:
		sheet = pe.get_sheet(file_name=filePath, name_columns_by_row = titleRow)
	except NotImplementedError as error:
		return {'error': 'No source found for ' + filePath}
	except FileNotFoundError as error:
		return {'error': error.args[1] + ': ' + filePath}

	titleRow = -1
	productInfoC = 0
	productInfoR = 0
	emptyRow = 1
	titles = ['Ref Des', 'Part Number', 'X-location', 'Y-location', 'Rotation', 'x']
	
	emptyRowList = sheet.row[emptyRow]
	if not all(isinstance(e, str) for e in emptyRowList):
		return{'error': 'File format incorrect: ' + filePath}
	spaceString = ''.join(emptyRowList)
	# for empty in emptyRowList:
	# 	spaceString = spaceString + empty
	if type(spaceString) is not str or spaceString.strip() != '':
		return{'error': 'File format incorrect: ' + filePath}

	try:
		productionInfo = sheet.row[productInfoR][productInfoC].split( )
		productName = productionInfo
		del sheet.row[emptyRow]
		del sheet.row[productInfoR]

		dataDictionary = sheet.to_dict()

		keys = list(dataDictionary.keys())
		for key in keys:
			if key[0] == ' ' or key[-1] == ' ':
				dataDictionary[key.strip()] = dataDictionary.pop(key)
	except:
		return {'error': 'File format incorrect: ' + filePath}

	missing = []
	for title in titles:
		if title not in dataDictionary:
			missing.append(title)

	if missing:
		return dataDictionary
	else:
		return {'error': missing}





