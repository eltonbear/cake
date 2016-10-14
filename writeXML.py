from xml.etree.ElementTree import ElementTree, Element, SubElement

def writeXml(version, excelData, XMLpath):
	partIndex = 2
	root = Element('Rec_SubstrateRecord')
	versionElement = SubElement(root, 'Version')
	versionElement.text = version
	alignmentNameElement = SubElement(root, 'AlignmentName')
	alignmentNameElement.text = excelData['productName']
	parts = excelData['data']
	for part in parts:
		partElement = writePart()
		root.insert(partIndex, partElement)

	tree = ElementTree(root)
	tree.write(XMLpath)

def writePart():
	partElement = Element('Rec_SubstratePlacement')
	enableElment = SubElement(partElement, 'Enabled')
	enableElment.text = '1'
	cameraNbrElement = SubElement(partElement, 'CameraNbr')
	cameraNbrElement.text = '1'
	pxElement = SubElement(partElement, 'PlacementCoords-X')
	pxElement.text = '0.55'
	pyElement = SubElement(partElement, 'PlacementCoords-Y')
	pyElement.text = '5'
	pAngleElement = SubElement(partElement, 'PlacementAngle')
	pAngleElement.text = '50'
	dieNameElement = SubElement(partElement, 'DieName')
	dieNameElement.text = 'asdfgh'
	cameraHeightElement = SubElement(partElement, 'CameraHeight')
	cameraHeightElement.text = '0'
	pForceElement = SubElement(partElement, 'PlacementForce')
	pForceElement.text = '0'
	referenceNumberElement = SubElement(partElement, 'Comment')
	referenceNumberElement.text = 'U3'
	dwellTimeElement = SubElement(partElement, 'DwellTime')
	dwellTimeElement.text = '0'
	probeSpeedElement = SubElement(partElement, 'ProbeSpeed')
	probeSpeedElement.text = '0'
	pUpSpeedElement = SubElement(partElement, 'PlacementUpSpeed')
	pUpSpeedElement.text = '0'
	airPuffTimerElement = SubElement(partElement, 'AirPuffTimer')
	airPuffTimerElement.text = '0'

	return partElement

writeXml('elton', 'N90406', 3)













































































































































































































































































































