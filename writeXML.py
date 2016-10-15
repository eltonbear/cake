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
	for part in parts:
		dieName = part['Part Number'].replace('-', '') + productName[-3:]
		partElement = writePart(part)
		root.insert(partIndex, partElement)

	tree = ElementTree(root)
	tree.write(XMLpath)

def writePart(partInfo):
	partElement = Element('Rec_SubstratePlacement')
	enableElment = SubElement(partElement, 'Enabled')
	enableElment.text = '1'
	cameraNbrElement = SubElement(partElement, 'CameraNbr')
	cameraNbrElement.text = '1'
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
def getFileNameInFolder(directory):
	print('in progress')

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
















































































































































































































































































































