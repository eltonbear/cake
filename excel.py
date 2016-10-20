import pyexcel as pe
import os

def readSheet(filePath):
	titleRow = 1
	fileBaseName = os.path.basename(filePath)
	fileName, fileExtension = os.path.splitext(fileBaseName)
	if fileExtension == '.csv':
		return {'error': '.csv file type does not support multiple sheets. Please save the file in .xlsx.'}
	try:
		# Get workbook that has two sheets
		book = pe.get_book(file_name=filePath) 
		if book.number_of_sheets() < 2:
			if 'Alignment' not in book.sheet_names():
				return {'error': 'File: ' + filePath + ' missing "Alignment" worksheet.'}
			if fileName not in book.sheet_names():
				return {'error': 'File: ' + filePath + ' missing "' + fileName + '" worksheet.'}

			# name_columns_by_row = titleRow)
	except NotImplementedError as error:
		return {'error': 'No source found for ' + filePath}
	except FileNotFoundError as error:
		return {'error': error.args[1] + ': ' + filePath}
		
	partSheet = book[fileName]
	alignmentSheet = book['Alignment']


	# titleRow = -1
	# productInfoC = 0
	# productInfoR = 0
	# emptyRow = 1
	# requireTitles = ['Ref Des', 'Part Number', 'X-location', 'Y-location', 'Rotation']
	# sheet.colnames = list(map(str.strip, sheet.colnames))
	# titles = sheet.colnames
	# emptyRowList = sheet.row[emptyRow]
	# missingColumns = []

	# # Check if require columns all exist
	# for requireTitle in requireTitles:
	# 	if requireTitle not in titles:
	# 		missingColumns.append(requireTitle)
	# if missingColumns:	
	# 	return{'error': 'File missing column of "' + ','.join(missingColumns) + '"!'}
	# # Check if the empty row is a list of strings
	# if not all(isinstance(e, str) for e in emptyRowList):
	# 	return{'error': 'File format incorrect: ' + filePath}
	# # Check if there an empty at row 2
	# spaceString = ''.join(emptyRowList)
	# if type(spaceString) is not str or spaceString.strip() != '':
	# 	return{'error': 'File format incorrect: ' + filePath}

	# try:
	# 	productionInfo = sheet.row[productInfoR][productInfoC].split( )
	# 	productName = productionInfo[0]
	# 	del sheet.row[emptyRow]
	# 	del sheet.row[productInfoR]
	# 	dataDictionary = sheet.to_records()
	# except:
	# 	return {'error': 'File format incorrect: ' + filePath}

	# return {'productName': productName, 'data': dataDictionary}

if __name__ == "__main__":
    data = readSheet(r'C:\Users\eltoshon\Desktop\cakeXMLfile\pico_top_expanded.xlsx')
    # print(data['productName'])
    # print(data['data'][0:2])




