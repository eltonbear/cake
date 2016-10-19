import os
import os.path
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import ElementTree, Element, SubElement

def writeXml(version, excelData,dieFolderPath, XMLpath):
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
	missingDieFile = []
	errorInFile = []
	error = False
	for part in parts:
		if part['Ref Des']:
			dieName = part['Part Number'].replace('-', '') + '_' + productName[-4:]
			if dieName not in dieToCameraNum:
				fileName = ''
				if dieName + '.XML' in dieFileList:
					fileName = dieName + '.XML'
				elif dieName + 'txt' in dieFileList:
					fileName = dieName + '.txt'
				else:
					missingDieFile.append(dieName)
					error = True
					###### are there supposed to be all die files i need?
				if fileName:
					xmlPath = dieFolderPath + '/' + fileName
					camNum = getCameraNum(xmlPath)
					if camNum.isdigit(): # if camNum is str, then there is error in getCameraNum
						dieToCameraNum[dieName] = camNum					
					else:
						errorInFile.append(camNum)
						error = True
			if not error:
				partElement = writePart(part, dieName, dieToCameraNum[dieName])
				root.insert(partIndex, partElement)
			##### MAYBE PUT LOCAL ALIGNMENT HERE
	if error:
		return {'missingDie': missingDieFile, 'error': errorInFile}
	else:
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
















































































































































































































































































































