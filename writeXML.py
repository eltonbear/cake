import os
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import ElementTree, Element, SubElement

def writeXml(version, excelData, XMLpath, dieFolderPath):
	partIndex = 2
	root = Element('Rec_SubstrateRecord')
	versionElement = SubElement(root, 'Version')
	versionElement.text = version
	productName = excelData['productName']
	alignmentNameElement = SubElement(root, 'AlignmentName')
	alignmentNameElement.text = productName
	parts = excelData['data']
	dieFileList = os.listdir(dieFolderPath)
	dieToCameraNum = {}
	for part in parts:
		dieName = part['Part Number'].replace('-', '') + productName[-3:]
		if dieName not in dieToCameraNum:
			fileNamesWanted = [dieName + '.xml', dieName + '.txt']
			camNum = '-1'
			for name in fileNamesWanted: # usually iterates 2 times
				if name in dieFileList:
					camNum = getCameraNum(dieFolderPath + '/' + name)
					dieToCameraNum{dieName: camNum}
					if not camNum.isdigit(): # if camNum is str, then there is error in getCameraNum
						return camNum
					break
			if camNum == '-1': # if camNum is still '-1', then the wanted die file is not in the folder
				return 'Missing die file for die: ' + dieName ###### are there supposed to be all die files i need?
		partElement = writePart(part, dieToCameraNum[dieName])
		root.insert(partIndex, partElement)
		##### MAYBE PUT LOCAL ALIGNMENT HERE

	tree = ElementTree(root)
	tree.write(XMLpath)
	return ''

def writePart(partInfo, cameraNbr):
	partElement = Element('Rec_SubstratePlacement')
	enableElment = SubElement(partElement, 'Enabled')
	enableElment.text = '1'
	cameraNbrElement = SubElement(partElement, 'CameraNbr')
	cameraNbrElement.text = cameraNbr
	pxElement = SubElement(partElement, 'PlacementCoords-X')
	pxElement.text = partInfo['X-location']
	pyElement = SubElement(partElement, 'PlacementCoords-Y')
	pyElement.text = partInfo['Y-location']
	pAngleElement = SubElement(partElement, 'PlacementAngle')
	pAngleElement.text = partInfo['Rotation']
	dieNameElement = SubElement(partElement, 'DieName')
	dieNameElement.text = partInfo['Part Number']
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

def getCameraNum(filePath):
	try:
		# Parse xml content
		tree = ET.parse(filePath)                                    
	except ET.ParseError as error: 
		return error.args[0] + ' in file: ' + filePath
	# Get root of the xml data tree and all reference and wire elements if xml is parsed successfully
	recDie = tree.getroot() 
	camerNumElement = recDie.find('CameraNbr')
	if camerNumElement != None:
		return camerNumElement.text
	else:
		return 'No camera number in file: ' + filePath

a = getCameraNum(r'C:\Users\eltoshon\Desktop\cakeXMLfile\1GG6400_7740.xml')
# b = getCameraNum(r'C:\Users\eltoshon\Desktop\cakeXMLfile\1GG6400_7740corrupt.xml')

print(a)
















































































































































































































































































































