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

		self.destroy()

	def getPaths(self):
		return {frameName: frame.getPath() for frameName, frame in self.frames.items()}

	def popErrorMessage(self, message):
		"""Show error message in a pop up window.

			Parameters
			----------
			message: string
				An error message.
		"""

		messagebox.showinfo("Warning", message, parent = self)

class browse(Frame):

	def __init__(self, parent, controller, nextFrame = None, prevFrame = None, message = None):
		"""Create a browse interface."""

		# Create a main frame
		Frame.__init__(self, parent, width = 1000)
		self.parent = parent		# Window for interface
		self.controller = controller
		self.filePath = ""			# Final file path
		self.filePathEntry = None	# File path in browse entry
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

		bCancel = Button(self, text = 'Cancel', width = 10, borderwidth = buttonBD, command = self.controller.closeWindow)	
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

	def getPath(self):

		return self.filePath

class browseSheet(browse):

	def __init__(self, parent, controller):
		message = 'Select an Excelsheet'
		browse.__init__(self,
		 parent, controller, 'browseFolder', message = message)

	def initGUI(self):
		super().initGUI()
		button = [{'text': 'Next', 'width': 7, 'function': self.GetEntrypathAndNextPage}]
		self.makeButtons(button)

	def getFilePathToEntry(self):
		fileType1 = ("Comma Delimited CSV", "*.CSV")
		fileType2 = ("Excel Workbook", "*.xlsx")
		fileType3 = ("Excel Macro-Enabled Workbook", "*.xlsm")
		fileType4 = ("All files", "*.*")

		path = askopenfilename(filetypes = (fileType1, fileType2, fileType3, fileType4), parent = self.parent)

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
		message = 'Save XML file'
		browse.__init__(self, parent, controller, prevFrame = 'browseFolder', message = message)

	def initGUI(self):
		super().initGUI()
		self.checkVar = IntVar()
		bCheck = Checkbutton(self, text = "Open XML", variable = self.checkVar, onvalue = 1, offvalue = 0, borderwidth = 2)
		bCheck.pack(side = LEFT, anchor=S, padx=6, pady=4)
		buttons = [{'text': 'Save', 'width': 7, 'function': self.save}, {'text': 'Back', 'width': 7, 'function': self.back}]
		self.makeButtons(buttons)

	def getFilePathToEntry(self):
		fileType1 = ("XML Data", "*.xml")
		fileType2 = ("Text", "*.txt")
		fileType3 = ("All files", "*.*")

		path = asksaveasfilename(defaultextension= '.xml', filetypes = (fileType1, fileType2, fileType3), parent = self.parent)

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

if __name__ == "__main__":
    app = cakeApp()
    app.mainloop()
 
