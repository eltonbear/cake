import os
import re
import math
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import ElementTree, Element, SubElement

def writeXml(excelData, dieFolderPath, XMLSaveFolderPath): 
	# Get the product name
	productName = excelData['productName'].replace('-', '')
	# Get last few digits of the product name
	lastFewDigits = productName[-4:]
	# Get a list of dictionaries of parts
	parts = excelData['data']
	# Get fiducials points
	fiducials = excelData['fiducials']  # fiducials = {'L1': {'p1': (x1, y1), 'p2': (x2, y2)}, 'L2': {'p1': (x1, y1), 'p2': (x2, y2)}}
	# Get a list of die file names in the die folder
	dieFileDict = {fileName.lower(): fileName for fileName in os.listdir(dieFolderPath)}
	# print(dieFileDict)
	# Added axes rotational angle into L1, if there is L1
	if 'L1' in fiducials:
		fiducials['L1']['axesRotation'] = calcAxesRotaionalAngle(fiducials['L1']['p1'], fiducials['L1']['p2'])
	# Added axes rotational angle into L2, if there is L2
	if 'L2' in fiducials:
		fiducials['L2']['axesRotation'] = calcAxesRotaionalAngle(fiducials['L2']['p1'], fiducials['L2']['p2'])
	
	refDesToPartDict = {}              		# {'Ref Des': partDictionary}
	layersToTipsToPartRootsDict = {}		# {'layer':{'tip number': fake root}}
	layersToLocalsToTipsToPartRootsDict = {}	# {'local1':{'tip number': fake root}, 'local2':{'tip number': fake root}}
	dieNameToTipAndCameraNumDict= {}		# {'dieName': {'camera number': camNum, 'tip number': tipNum}}
	missingDieFile = set()             		# A set of missing die file names if there any
	errorInFile = []						# A list of error messages when the process is run if there is any
	error = False							# Starts with no errors
	# Iterate through all parts dictionaries in the list
	for part in parts:
		# Get Ref Des value, and If 'Ref Des' value exists in the dictionary, do the following
		refDes = part['Ref Des']
		if refDes:
			# Store parts in dictionary
			refDesToPartDict[refDes] = part
			# A die name is part number plus an underscore and the last four digit of the product name
			dieName = part['Part Number'].replace('-', '') + '_' + lastFewDigits
			# Grab the corresponding tip number and camera number to die name from the die xml file if not already looked up
			if dieName not in dieNameToTipAndCameraNumDict:
				# There are two possible file extension for a xml file (.txt and .xml). If both of these file paths not exist in the folder, gives an error.
				dieFileName = ''
				dieNameLower = dieName.lower()
				dieFileNameLower = dieNameLower + '.xml'
				if dieFileNameLower in dieFileDict:
					dieFileName = dieFileDict[dieFileNameLower]
				elif dieNameLower + 'txt' in dieFileDict:
					dieFileName = dieFileDict[dieNameLower + '.txt']
				else:
					missingDieFile.add(dieName)
					error = True
				# If file path is found, access the file and get the tip number and camera number. Gives errors if there are exceptions parsing the xml file or finding the desire elements
				if dieFileName:
					dieFilePath = dieFolderPath + '/' + dieFileName
					TipAndCamNum = getTipAndCameraNum(dieFilePath)
					if 'error' in TipAndCamNum: # if there is 'error' key in TipAndCamNum, then there is error in getCameraNum
						errorInFile.append(TipAndCamNum['error'])
						error = True					
					else:
						dieNameToTipAndCameraNumDict[dieName] = TipAndCamNum
			# If there isn't any errors throughout the previous processes, write a part element and append it to the corresponding root based on its tip number.
			if not error:
				# Get layer name from dictionary
				layer = part['Layer']
				# Get tip number from dictionary
				tipNumber = dieNameToTipAndCameraNumDict[dieName]['TipNbr']
				# Get layer name and base alignment if it's not either L1' or 'L2'
				if layer != 'L1' and layer != 'L2':
					local, layer = layer.split('_')
					if layer not in layersToLocalsToTipsToPartRootsDict:
						layersToLocalsToTipsToPartRootsDict[layer] = {}
					# Write a part element with given imformation
					partE = writeLocalAlignmentPart(part, dieName, dieNameToTipAndCameraNumDict[dieName]['CameraNbr'], fiducials, layer)
					createFakeRoot(layersToLocalsToTipsToPartRootsDict[layer], local, tipNumber, partE)
				else:
					partE = writePart(part, dieName, dieNameToTipAndCameraNumDict[dieName]['CameraNbr'], fiducials)
					# Write a part element with given imformation
					createFakeRoot(layersToTipsToPartRootsDict, layer, tipNumber, partE)

			error = False
	# print(dieNameToTipAndCameraNumDict, '\n')
	# print(layerToBaseAlignment, '\n')
	# print(layersToTipsToPartRootsDict, '\n')
	# print(layersToLocalsToTipsToPartRootsDict)

	# Combine roots to one root in the order of tip number and write them into seperate files
	for layer, roots in layersToTipsToPartRootsDict.items():
		finalRoot = creatTemplate(layer, productName)
		combineRoots(finalRoot, productName, roots)
		tree = ElementTree(finalRoot)
		path = XMLSaveFolderPath + '/' + productName + '_' + layer + '.XML'
		tree.write(path)

	
	for layer, localss in layersToLocalsToTipsToPartRootsDict.items():
		finalRoot = creatTemplate(layer, productName)
		localKeys = sorted(localss.keys(), key = naturalKey)
		for local in localKeys:
			finalRoot.append(writeLocalOpen(local + '_' + lastFewDigits))
			combineRoots(finalRoot, productName, localss[local])
			finalRoot.append(writeLocalClosed(local + '_' + lastFewDigits))

		tree = ElementTree(finalRoot)
		path = XMLSaveFolderPath + '/' + productName + '_Local_' + layer + '.XML'
		tree.write(path)

	# If there are any errors, return the error messaegs
	if missingDieFile or errorInFile:
		return {'missingDie': list(missingDieFile), 'error': errorInFile}
	else:
		return {}

def createFakeRoot(layerToTipToRoots, layer, tipNumber, partElement):
	if layer in layerToTipToRoots:
		# Get the dictionary of tip numbers to part elements for that one layer
		partsOnOneLayer = layerToTipToRoots[layer]
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
		layerToTipToRoots[layer] = {tipNumber: fakeRoot}

	return layerToTipToRoots

def creatTemplate(layer, productName):
	# Create a root of the xml file
	root = Element('Rec_SubstrateRecord')
	# Create a version element and its content
	versionElement = SubElement(root, 'Version')
	versionElement.text = '7.6.22'
	# Create an alignment name element with a content of the product name
	alignmentNameElement = SubElement(root, 'AlignmentName')
	alignmentNameElement.text = productName + '_' + layer

	return root

def combineRoots(root, productName, fakeRootsSortedWithTipNums):

	sortedTipNumber = sorted(fakeRootsSortedWithTipNums.keys(), key = int)
	for tipSize in sortedTipNumber:
		root.extend(fakeRootsSortedWithTipNums[tipSize])

def writePartCommon(partElement, partInfo, dieName, cameraNbr, fiducials, layer):
	enableElment = SubElement(partElement, 'Enabled')
	enableElment.text = '1'
	cameraNbrElement = SubElement(partElement, 'CameraNbr')
	cameraNbrElement.text = cameraNbr
	pxElement = SubElement(partElement, 'PlacementCoords-X')
	x, y = pointTranslation((partInfo['X-location'], partInfo['Y-location']), fiducials[layer]['p1'],  fiducials[layer]['axesRotation'])
	pxElement.text = str(x)
	pyElement = SubElement(partElement, 'PlacementCoords-Y')
	pyElement.text = str(y)
	pAngleElement = SubElement(partElement, 'PlacementAngle')
	pAngleElement.text = str(calcPartAngle(math.radians(partInfo['Rotation']), fiducials[layer]['axesRotation']))
	dieNameElement = SubElement(partElement, 'DieName')
	dieNameElement.text = dieName

def writePart(partInfo, dieName, cameraNbr, fiducials):
	partElement = Element('Rec_SubstratePlacement')
	writePartCommon(partElement, partInfo, dieName, cameraNbr, fiducials, partInfo['Layer'])
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

def writeLocalAlignmentPart(partInfo, dieName, cameraNbr, fiducials, layer):
	partElement = Element('Rec_SubstratePlacement')	
	writePartCommon(partElement, partInfo, dieName, cameraNbr, fiducials, layer)
	referenceNumberElement = SubElement(partElement, 'Comment')
	referenceNumberElement.text = partInfo['Ref Des']

	return partElement
def writeLocalOpen(alignmentName):
	openElement = Element('Rec_SubstrateLocalAlignment')
	enableElment = SubElement(openElement, 'Enabled')
	enableElment.text = '1'
	alignmentNameElement = SubElement(openElement, 'AlignmentName')
	alignmentNameElement.text = alignmentName

	return openElement

def writeLocalClosed(alignmentName):
	closeElement = Element('Rec_SubstrateSwitchOff')
	enableElment = SubElement(closeElement, 'Enabled')
	enableElment.text = '1'
	numberLevelElement = SubElement(closeElement, 'NumberLevel')
	numberLevelElement.text = '1'
	alignmentNameElement = SubElement(closeElement, 'AlignmentName')
	alignmentNameElement.text = alignmentName

	return closeElement

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


def calcAxesRotaionalAngle(p1, p2):
	"""Calculate how much the new axes rotate in raidan. Clockwise rotation results a negative angle.

		Parameters
			----------
			p1: tuple
				The first fiducial point, which is also the new origin. p1 --> (x ,y).
			p2: tiple
				The second fiducial point. p2 --> (x, y)

		Return: angle in radian
	"""
	deltaX = p2[0]-p1[0]
	deltaY = p2[1]-p1[1]

	return math.atan2(deltaY, deltaX)

def calcPartAngle(ogAngle, axesRotationalAngle):
	"""Calculate the acutaly angle of a part in the new coordinate system.

		Parameters
			----------
			ogAngle: int
				Original angle the part is pointing in radian.
			axesRotationalAngle: float
				The angle that the axes rotated in radian.

		Return: 
			----------
			finalAngle: float
				The new angle that a part points to in the new system in radian.
	"""

	finalAngle = (ogAngle - axesRotationalAngle) % (2*math.pi)
	if finalAngle > math.pi:
		finalAngle = finalAngle - (2*math.pi)
	return -finalAngle

def pointTranslation(p, p0, axesRotationalAngle):
	"""Translate a point to a new coordinate in the new coordinate system. Note that the new coordinate is fliped. 
		Positive Y points down, so and negative sign is added in front of returned y value.

		Parameters
			----------
			p: tuple
				A point to be translated. p --> (x, y)
			p0: tiple
				The new origin of the new coordinate system. p0 --> (x, y)
			axesRotationalAngle: float
				An angle of axes rotation in radian 

		Return: x and y
	"""
	radius = math.sqrt((p0[0]-p[0])**2+(p0[1]-p[1])**2)
	# The angle to the point is equal to the angle between new origin to the point and new x axis after axes are shift but not rotated minus axesRotationalAngle 
	pointAngle = calcAxesRotaionalAngle(p0, p) - axesRotationalAngle
	newX = math.cos(pointAngle) * radius
	newY = math.sin(pointAngle) * radius
	return round(newX, 8), -round(newY, 8)

def naturalKey(string):
    """See http://www.codinghorror.com/blog/archives/001018.html"""
    return [int(s) if s.isdigit() else s for s in re.split(r'(\d+)', string)]

def calcAxesRotaionalAngleTest():

	print(math.degrees(calcAxesRotaionalAngle((0, 0), (5, 5))))
	print(math.degrees(calcAxesRotaionalAngle((0, 0), (0, 5))))
	print(math.degrees(calcAxesRotaionalAngle((0, 0), (-5, 5))))
	print(math.degrees(calcAxesRotaionalAngle((0, 0), (-5, 0))))
	print(math.degrees(calcAxesRotaionalAngle((0, 0), (-5, -5))))
	print(math.degrees(calcAxesRotaionalAngle((0, 0), (0, -5))))
	print(math.degrees(calcAxesRotaionalAngle((0, 0), (5, -5))))

def calcPartAngleTest():

	print(math.degrees(calcPartAngle(0, calcAxesRotaionalAngle((0, 0), (5, 5)))))
	print(math.degrees(calcPartAngle(math.pi/2, calcAxesRotaionalAngle((0, 0), (5, 5)))))
	print(math.degrees(calcPartAngle(math.pi, calcAxesRotaionalAngle((0, 0), (5, 5)))))
	print(math.degrees(calcPartAngle(math.pi/2*3, calcAxesRotaionalAngle((0, 0), (5, 5)))))
	print('\n')
	print(math.degrees(calcPartAngle(0, calcAxesRotaionalAngle((0, 0), (-5, 5)))))
	print(math.degrees(calcPartAngle(math.pi/2, calcAxesRotaionalAngle((0, 0), (-5, 5)))))
	print(math.degrees(calcPartAngle(math.pi, calcAxesRotaionalAngle((0, 0), (-5, 5)))))
	print(math.degrees(calcPartAngle(math.pi/2*3, calcAxesRotaionalAngle((0, 0), (-5, 5)))))
	print('\n')
	print(math.degrees(calcPartAngle(0, calcAxesRotaionalAngle((0, 0), (5, -5)))))
	print(math.degrees(calcPartAngle(math.pi/2, calcAxesRotaionalAngle((0, 0), (5, -5)))))
	print(math.degrees(calcPartAngle(math.pi, calcAxesRotaionalAngle((0, 0), (5, -5)))))
	print(math.degrees(calcPartAngle(math.pi/2*3, calcAxesRotaionalAngle((0, 0), (5, -5)))))

def pointTranslationTest():

	print(pointTranslation((2, -2), (5, 5), calcAxesRotaionalAngle((5, 5), (0, 10))))
	print(pointTranslation((-2, -2), (5, 5), calcAxesRotaionalAngle((5, 5), (0, 10))))
	print(pointTranslation((2, -2), (5, 5), calcAxesRotaionalAngle((5, 5), (10, 10))))
	print(pointTranslation((2, -2), (-2, -2), calcAxesRotaionalAngle((-2, -2), (-5, -5))))

# if __name__ == '__main__':
# 	# import excel
# 	# eee = excel.readSheet(r'C:\Users\eltoshon\Desktop\cakeXMLfile\pico_top_expanded15.xlsx')
# 	# e = writeXml(eee, r'C:\Users\eltoshon\Desktop\Die', r'C:\Users\eltoshon\Desktop\cakeXMLfile')
# 	# print(e)

# 	# calcAxesRotaionalAngleTest()
# 	# print('\n')
# 	# calcPartAngleTest()
# 	# print('\n')
# 	# pointTranslationTest()
# 	print(calcAxesRotaionalAngle((0.70663018,	1.15696974), (0.23344307, 0.80796973)),math.degrees(calcAxesRotaionalAngle((0.70663018,	1.15696974), (0.23344307, 0.80796973))) ,'calcAxesRotaionalAngle')

# 	print(calcPartAngle(math.pi*3/2, calcAxesRotaionalAngle((0.70663018,	1.15696974), (0.23344307, 0.80796973))), math.degrees(calcPartAngle(math.pi*3/2, calcAxesRotaionalAngle((0.70663018,	1.15696974), (0.23344307, 0.80796973)))), 'calcPartAngle')


	# local alignment angle???????????????????
	# adjust error message configuration