import interface
import excel
import writeXML
import os

def runApp():
	app = interface.cakeApp()
	app.mainloop()
	paths = app.getPaths()
	openXML = app.ifOpenXML()
	# print(paths, app.toContinue())
	if app.toContinue():
		partsDataFromExcel = excel.readSheet(paths['browseSheet'])
		print(55)
		if 'error' in partsDataFromExcel:
			app.errorMessageWindow(partsDataFromExcel['error'])
			print(44)
		else:
			error = writeXML.writeXml('version??', partsDataFromExcel, paths['browseFolder'], paths['saveXML'])
			print(33)
			if error:
				if error['missingDie'] and error['error'] :
					message =  '\n'.join(['Missing Die Files: '] + error['missingDie'] + ['\n', 'Error'] + error['error'])
				elif error['missingDie']:
					message =  '\n'.join(['Missing Die Files: '] + error['missingDie'])
				else:
					message =  '\n'.join(['Error'] + error['error'])
				app.errorMessageWindow(message)
				print(22)
			if openXML:
				print(1111)
				os.startfile(paths['saveXML'])

if __name__ == "__main__":
	runApp()
    