from tkinter import *
from tkinter.filedialog import askopenfilename, askdirectory, asksaveasfilename
from tkinter import messagebox
from os.path import exists

class cakeApp(Tk):

	def __init__(self, *args, **kwargs):
		Tk.__init__(self, *args, **kwargs)
		Tk.wm_title(self, "MRSI Cake")
		container = Frame(self)
		container.pack(side="top", fill="both", expand=True)
		container.grid_rowconfigure(0, weight=1)            
		container.grid_columnconfigure(0, weight=1)

		self.frames = {}
		for frameClass in (browseSheet, browseFolder, saveXML):
			pageName = frameClass.__name__
			frame = frameClass(parent = container, controller = self)
			self.frames[pageName] = frame

			# put all of the pages in the same location;
			# the one on the top of the stacking order
			# will be the one that is visible.

		self.showFrame("browseSheet")

	def getFrame(self, frameName):

		return self.frames[frameName]

	def showFrame(self, pageName):
		'''Show a frame for the given page name'''

		frame = self.frames[pageName]
		frame.tkraise()

	def closeWindow(self):
		'''Close the window.'''
		self.isCancelled = True
		self.destroy()

	def toContinue(self):
		return all([frame.toContinue() for frame in self.frames.values()])

	def getPaths(self):
		print(self.frames['saveXML'].getPath())
		return {frameName: frame.getPath() for frameName, frame in self.frames.items()}

	def ifOpenXML(self):

		return self.getFrame('saveXML').getCheckBoxVal()

	def popErrorMessage(self, message):
		"""Show error message in a pop up window.

			Parameters
			----------
			message: string
				An error message.
		"""

		messagebox.showinfo("Warning", message, parent = self)

	def errorMessageWindow(self, message):
		window = Tk()
		errorMessage(window, message)
		window.mainloop()


class browse(Frame):

	def __init__(self, parent, controller, nextFrame = None, prevFrame = None, message = None):
		"""Create a browse interface."""

		# Create a main frame
		Frame.__init__(self, parent, width = 1000)
		self.parent = parent		# Window for interface
		self.controller = controller
		self.filePath = ""			# Final file path
		self.filePathEntry = None	# File path in browse entry
		self.isCancelled = False
		self.nextFrame = nextFrame
		self.prevFrame = prevFrame
		self.message = message
		# Create GUI
		self.initGUI()

	def initGUI(self):
		"""Create GUI including buttons and path entry."""

		# Set main frame's location 
		self.grid(row=0, column=0, sticky="nsew")

		# Set path entry frame and its location
		self.entryFrame = Frame(self, relief = RAISED, borderwidth = 1)
		self.entryFrame.pack(fill = BOTH, expand = False)
		# Make label
		if self.message:
			messageLabel = Label(self.entryFrame, text = self.message, font=("Bradley", 10))
			messageLabel.pack(anchor=W,  padx=0, pady=0)

		# Set path entry and its location
		self.filePathEntry = Entry(self.entryFrame, bd = 4, width = 50)
		self.filePathEntry.pack(side = LEFT,  padx=2, pady=1)


	def makeButtons(self, buttonInfo):
		"""Create buttons."""
		lowerButtonPadx = 4
		lowerButtonPady = 3
		buttonBD = 3

		# Set browse button and location
		bBrowse = Button(self.entryFrame, text = 'Browse', width = 10, borderwidth = buttonBD, command = self.getFilePathToEntry)	
		bBrowse.pack(side = LEFT, padx=4, pady=2)

		bCancel = Button(self, text = 'Cancel', width = 10, borderwidth = buttonBD, command = self.closeWindow)	
		bCancel = bCancel.pack(side = RIGHT, anchor = S, padx=lowerButtonPadx, pady=lowerButtonPady)
		for button in buttonInfo:
			# Set cancel button and location
			b = Button(self, text = button['text'], width = button['width'], borderwidth = buttonBD, command = button['function'])
			b.pack(side = RIGHT, anchor = S, padx=lowerButtonPadx, pady=lowerButtonPady)

	def getFilePathToEntry(self):
		"""Get file Path to file path entry."""

		path = askopenfilename(filetypes = ("All files", "*.*"), parent = self.parent)

		# Once self.filePath gets a filepath, delete what's in the entry and put self.filePath into the entry
		if path != '':
			self.filePathEntry.delete(0, 'end')
			self.filePathEntry.insert(0, path)

	def GetEntrypathAndNextPage(self):
		"""A set of instructions are ran when ok button is clicked."""

		# Get a file path from browe entry
		self.filePath = self.filePathEntry.get()
		# When entry is empty						
		if self.filePath == '':
			self.controller.popErrorMessage('Please select a file or folder!')
		# When file path is invaild meaning not a legit file
		elif not exists(self.filePath):
			self.controller.popErrorMessage('Path does not exist!')
		else:
			self.controller.showFrame(self.nextFrame)

	def back(self):
		"""Go back to the previous frame (page)."""

		self.controller.showFrame(self.prevFrame)

	def closeWindow(self):
		self.isCancelled = True
		self.controller.closeWindow()

	def toContinue(self):
		return not self.isCancelled

	def getPath(self):

		return self.filePath

class browseSheet(browse):

	def __init__(self, parent, controller):
		message = 'Select an Excelsheet (No .csv file)'
		browse.__init__(self, parent, controller, 'browseFolder', message = message)

	def initGUI(self):
		super().initGUI()
		button = [{'text': 'Next', 'width': 7, 'function': self.GetEntrypathAndNextPage}]
		self.makeButtons(button)

	def getFilePathToEntry(self):
		fileType1 = ("Excel Workbook", "*.xlsx")
		fileType2 = ("Excel Macro-Enabled Workbook", "*.xlsm")
		fileType3 = ("All files", "*.*")

		path = askopenfilename(filetypes = (fileType1, fileType2, fileType3), parent = self.parent)

		# Once self.filePath gets a filepath, delete what's in the entry and put self.filePath into the entry
		if path != '':
			self.filePathEntry.delete(0, 'end')
			self.filePathEntry.insert(0, path)

class browseFolder(browse):

	def __init__(self, parent, controller):
		message = 'Select a die folder'
		browse.__init__(self, parent, controller, 'saveXML', 'browseSheet', message)

	def initGUI(self):
		super().initGUI()
		buttons = [{'text': 'Next', 'width': 7, 'function': self.GetEntrypathAndNextPage}, {'text': 'Back', 'width': 7, 'function': self.back}]
		self.makeButtons(buttons)

	def getFilePathToEntry(self):
		
		path = askdirectory(parent = self.parent)

		# Once self.filePath gets a filepath, delete what's in the entry and put self.filePath into the entry
		if path != '':
			self.filePathEntry.delete(0, 'end')
			self.filePathEntry.insert(0, path)

class saveXML(browse):

	def __init__(self, parent, controller):
		message = 'Select a folder to save XML files'
		browse.__init__(self, parent, controller, prevFrame = 'browseFolder', message = message)

	def initGUI(self):
		super().initGUI()
		self.checkVar = IntVar()
		bCheck = Checkbutton(self, text = "Open XML", variable = self.checkVar, onvalue = 1, offvalue = 0, borderwidth = 2)
		bCheck.pack(side = LEFT, anchor=S, padx=6, pady=4)
		buttons = [{'text': 'Save', 'width': 7, 'function': self.save}, {'text': 'Back', 'width': 7, 'function': self.back}]
		self.makeButtons(buttons)

	def getFilePathToEntry(self):

		path = askdirectory(parent = self.parent)

		# Once self.filePath gets a filepath, delete what's in the entry and put self.filePath into the entry
		if path != '':
			self.filePathEntry.delete(0, 'end')
			self.filePathEntry.insert(0, path)

	def save(self):
		# Get a file path from browe entry
		self.filePath = self.filePathEntry.get()
		# When entry is empty						
		if self.filePath == '':
			self.controller.popErrorMessage('Please input file path!')
		else:
			self.controller.closeWindow()

	def getCheckBoxVal(self):
		return self.checkVar.get()
 
class errorMessage(Frame):
	"""A class of error message interface that presents saves(optional) error messages.

		Parameters
		----------
		parent: Tk
			A window for an application.
		message: string
			An error message.
		textFilePath: string
			An error message text file destination. 
	"""

	def __init__(self, parent, message):
		"""Creat an error message interface."""

		self.parent = parent				# Main window
		self.message = message 				# Error message
		# Creat GUI
		self.initGUI()

	def initGUI(self):
		"""Creat GUI including a title, message, and buttons."""

		# Set window's title
		self.parent.title("Error Message")
		# Creat frames that contain messages and buttons 
		self.buttonFrame = Frame(self.parent)
		self.buttonFrame.pack(fill = BOTH, expand = True)
		messageFrame = Frame(self.buttonFrame, borderwidth = 1)
		messageFrame.pack(fill = BOTH, expand = True)
		# Creat buttons
		self.makeButtons()
		# Create and show an error message as an label
		var = StringVar()
		label = Message(messageFrame, textvariable=var, relief=RAISED, width = 1000)
		var.set(self.message)
		label.pack(fill = BOTH, expand = True)

	def makeButtons(self):
		"""Create all buttons."""

		# Create save and ok buttons and set their locations
		bSave = Button(self.buttonFrame, text = "Save", width = 5, command = self.save)
		bSave.pack(side = RIGHT, padx=5, pady=2)
		bOk = Button(self.buttonFrame, text = "Ok", width = 5, command = self.parent.destroy)
		bOk.pack(side = RIGHT, padx=3, pady=2)

	def save(self):
		fileType1 = ("Text", "*.txt")
		fileType2 = ("All files", "*.*")
		path = asksaveasfilename(defaultextension= '.txt', filetypes = (fileType1, fileType2), parent = self.parent)
		self.writeToText(path)

	def writeToText(self, textFilePath):
		"""Write an error message into a text file and save it in the current directory."""

		# Create a text file for write only mode in the current directory
		file = open(textFilePath,'w')
		# Write a message into file 
		file.write(self.message)
		# Close file
		file.close()
		# Close the interface
		self.parent.destroy()

app = cakeApp()
app.mainloop()