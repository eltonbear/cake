import os
import os.path
import math
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import ElementTree, Element, SubElement

def writeXml(excelData,dieFolderPath, XMLpath):
	# Get the product name
	productName = excelData['productName']
	# Get a list of dictionaries of parts
	parts = excelData['data']
	# Create a root of the xml file
	root = Element('Rec_SubstrateRecord')
	# Create a version element and its content
	versionElement = SubElement(root, 'Version')
	versionElement.text = '7.6.22'
	# Create an alignment name element with a content of the product name
	alignmentNameElement = SubElement(root, 'AlignmentName')
	alignmentNameElement.text = productName
	# Get a list of die file names in the die folder
	dieFileList = os.listdir(dieFolderPath)

	refDesToPartDict = {}              	# {'Ref Des': partDictionary}
	layersToTipsToPartElementsDict = {} # {'layer':{'tip number': fake root}}
	dieNameToTipAndCameraNumDict= {}	# {'dieName': {'camera number': camNum, 'tip number': tipNum}}
	missingDieFile = []                 # A list of missing die file names if there any
	errorInFile = []					# A list of error messages when the process is run if there is any
	error = False						# Starts with no errors
	# Iterate through all parts dictionaries in the list
	for part in parts:
		# Get Ref Des value, and If 'Ref Des' value exists in the dictionary, do the following
		refDes = part['Ref Des']
		if refDes:
			# Store parts in dictionary
			refDesToPartDict[refDes] = part
			# A die name is part number plus an underscore and the last four digit of the product name
			dieName = part['Part Number'].replace('-', '') + '_' + productName[-4:]
			# Grab the corresponding tip number and camera number to die name from the die xml file if not already looked up
			if dieName not in dieNameToTipAndCameraNumDict:
				# There are two possible file extension for a xml file (.txt and .xml). If both of these file paths not exist in the folder, gives an error.
				fileName = ''
				if dieName + '.XML' in dieFileList:
					fileName = dieName + '.XML'
				elif dieName + 'txt' in dieFileList:
					fileName = dieName + '.txt'
				else:
					missingDieFile.append(dieName)
					error = True
					######################################################################## are there supposed to be all die files i need? 
				# If file path is found, access the file and get the tip number and camera number. Gives errors if there are exceptions parsing the xml file or finding the desire elements
				if fileName:
					xmlPath = dieFolderPath + '/' + fileName
					TipAndCamNum = getTipAndCameraNum(xmlPath)
					if 'error' in TipAndCamNum: # if there is 'error' key in TipAndCamNum, then there is error in getCameraNum
						errorInFile.append(TipAndCamNum['error'])
						error = True					
					else:
						dieNameToTipAndCameraNumDict[dieName] = TipAndCamNum
			# If there isn't any errors throughout the previous processes, write a part element and append it to root of the xml tree.
			if not error:
				# Get layer name from dictionary
				layer = part['Layer'] ######################################################### assume the title is called 'Layer'
				# Get tip number from dictionary
				tipNumber = TipAndCamNum['TipNbr']
				# Change the layer name to 'localAlignment' if it's not either 'PCB' or 'Miboard'
				if layer != 'PCB' or layer != 'Miboard':  ##################################### what are the names for layer
					layer = 'localAlignment'
				# Write a part element with given imformation
				partElement = writePart(part, dieName, TipAndCamNum['CameraNbr'])

				# If the layer ket is already in the dictionary				
				if layer in layersToTipsToPartElementsDict:
					# Get the dictionary of tip numbers to part elements for that one layer
					partsOnOneLayer = layersToTipsToPartElementsDict[layer]
					# If tip number key is in the dictionary
					if tipNumber in partsOnOneLayer:
						# Append the part to the root
						partsOnOneLayer[tipNumber].append(partElement)
					else:
						# Create a sudo root for the xml file
						fakeRoot = Element('Rec_SubstrateRecord')
						# Append the part to the root
						fakeRoot.append(partElement)
						# Asign the tip number key to the root in the dictionary
						partsOnOneLayer[tipNumber] = fakeRoot
				else:
					# Create a sudo root for the xml file
					fakeRoot = Element('Rec_SubstrateRecord')
					# Append the part to the root
					fakeRoot.append(partElement)
					# Asign the layer key to the tip number key to the root in the dictionary
					layersToTipsToPartElementsDict[layer] = {tipNumber: fakeRoot}
	# If there are any errors, return the error messaegs
	if error:
		return {'missingDie': missingDieFile, 'error': errorInFile}
	else:
		### combine trees useing extend. and save them into 3 different files
		tree = ElementTree(root)
		tree.write(XMLpath)
		return {}

def writePart(partInfo, dieName, cameraNbr):
	partElement = Element('Rec_SubstratePlacement')
	enableElment = SubElement(partElement, 'Enabled')
	enableElment.text = '1'
	cameraNbrElement = SubElement(partElement, 'CameraNbr')
	cameraNbrElement.text = cameraNbr
	pxElement = SubElement(partElement, 'PlacementCoords-X')
	pxElement.text = str(partInfo['X-location'])
	pyElement = SubElement(partElement, 'PlacementCoords-Y')
	pyElement.text = str(partInfo['Y-location'])
	pAngleElement = SubElement(partElement, 'PlacementAngle')
	pAngleElement.text = str(partInfo['Rotation'])
	dieNameElement = SubElement(partElement, 'DieName')
	dieNameElement.text = dieName
	cameraHeightElement = SubElement(partElement, 'CameraHeight')
	cameraHeightElement.text = '0'
	pForceElement = SubElement(partElement, 'PlacementForce')
	pForceElement.text = '0'
	referenceNumberElement = SubElement(partElement, 'Comment')
	referenceNumberElement.text = partInfo['Ref Des']
	dwellTimeElement = SubElement(partElement, 'DwellTime')
	dwellTimeElement.text = '0'
	probeSpeedElement = SubElement(partElement, 'ProbeSpeed')
	probeSpeedElement.text = '0'
	pUpSpeedElement = SubElement(partElement, 'PlacementUpSpeed')
	pUpSpeedElement.text = '0'
	airPuffTimerElement = SubElement(partElement, 'AirPuffTimer')
	airPuffTimerElement.text = '0'

	return partElement

def getTipAndCameraNum(filePath):
	try:
		# Parse xml content
		tree = ET.parse(filePath)                                    
	except ET.ParseError as error: 
		return {'error': error.args[0] + ' in file: ' + filePath}
	# Get root of the xml data tree and all reference and wire elements if xml is parsed successfully
	recDie = tree.getroot() 
	camerNumElement = recDie.find('CameraNbr')
	tipNumElement = recDie.find('TipNbr')
	if camerNumElement != None:
		if tipNumElement != None:
			return {'CameraNbr': camerNumElement.text, 'TipNbr': tipNumElement.text}
		else:
			return {'error': 'No tip number in file: ' + filePath}
	else:
		return {'error':'No camera number in file: ' + filePath}


def calcAxesRotaionalAngle(x1,y1,x2,y2):
	deltaX = x2-x1
	deltaY = y2-y1

	return math.atan2(deltaY, deltaX)

def calcPartAngle(ogAngle, axesRotationalAngle):
	
	finalAngle = (ogAngle - axesRotationalAngle) % (2*math.pi)
	if finalAngle > math.pi:
		finalAngle = finalAngle - (2*math.pi)
	print(math.degrees(axesRotationalAngle), math.degrees(finalAngle))	
	return round(finalAngle, 8)

def pointTranslation(x1, y1, xo, yo, axesRotationalAngle):
	radius = math.sqrt((xo-x1)**2+(yo-y1)**2)
	print(radius, ' r')
	# The angle to the point is equal to the angle between new origin to the point and new x axis after axes are shift but not rotated - axesRotationalAngle 
	pointAngle = calcAxesRotaionalAngle(xo, yo, x1, y1) - axesRotationalAngle
	print('calcAxesRotaionalAngle', math.degrees(calcAxesRotaionalAngle(xo, yo, x1, y1)))
	print('axesRotationalAngle', math.degrees(axesRotationalAngle))
	print('pointAngle', math.degrees(pointAngle))
	newX = math.cos(pointAngle) * radius
	newY = math.sin(pointAngle) * radius
	print(round(newX, 8), round(newY, 8))
	return round(newX, 8), -round(newY, 8)

if __name__ == '__main__':

	# print(calcPartAngle(math.pi/2, -110/180*math.pi))
	# a = calcAxesRotaionalAngle(5, 5 , 0, 10)
	# print(math.degrees(a))
	# b ,c  = pointTranslation(2, -2, 5, 5, a)
	# a = getTipAndCameraNum(r'C:\Users\eltoshon\Desktop\Die\1GC14038_7740.xml')
	# print(type(a['TipNbr']))
	# print(a['TipNbr'])


# how string will be in layer 3 (part name? ) also is this hard for the designer? 
# How is the angle presented? there are negative and positive values
# Rec_SubstrateSwitchOff?
# localAlightmet?