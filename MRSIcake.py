import interface
import excel
import writeXML
import os

def runApp():
	app = interface.cakeApp()
	app.mainloop()
	paths = app.getPaths()
	openXML = app.ifOpenXML()
	if app.toContinue():
		partsDataFromExcel = excel.readSheet(paths['browseSheet'])
		if 'error' in partsDataFromExcel:
			app.errorMessageWindow(partsDataFromExcel['error'])
		else:
			ifError = writeXML.writeXml(partsDataFromExcel, paths['browseFolder'], paths['saveXML'])
			if ifError:
				if ifError['missingDie'] and ifError['error'] :
					message = '\n'.join(['Missing Die Files: '] + ifError['missingDie'] + ['\n', 'Error'] + ifError['error'])
				elif ifError['missingDie']:
					message = '\n'.join(['Missing Die Files: '] + ifError['missingDie'])
				else:
					message = '\n'.join(['Error'] + ifError['error'])
				app.errorMessageWindow(message)
			if openXML:
				for xmlPath in ifError['path']:
					os.startfile(xmlPath)

if __name__ == "__main__":
	runApp()
    