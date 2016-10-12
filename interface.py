from tkinter import *
from tkinter.filedialog import askopenfilename, askdirectory, asksaveasfilename
from tkinter import messagebox
from os.path import exists

class browse(Frame):

	def __init__(self, root):
		"""Create a browse interface."""

		# Create a main frame
		Frame.__init__(self, root, width = 1000)
		self.root = root		# Window for interface
		self.filePath = ""			# Final file path
		self.filePathEntry = None	# File path in browse entry
		# Create GUI
		self.initGUI()

	def initGUI(self):
		"""Create GUI including buttons and path entry."""

		# Set main frame's location 
		self.pack(fill = BOTH, expand = True)
		# Set path entry frame and its location
		self.entryFrame = Frame(self, relief = RAISED, borderwidth = 1)
		self.entryFrame.pack(fill = BOTH, expand = True)
		# Set path entry and its location
		self.filePathEntry = Entry(self.entryFrame, bd = 4, width = 50)
		self.filePathEntry.grid(row = 0, column = 2, columnspan = 5, padx=2, pady=2)

	def makeButtons(self, getFilePath, buttonInfo):
		"""Create buttons."""

		# Set browse button and location
		bBrowse = Button(self.entryFrame, text = 'Browse', width = 10, command = getFilePath)	
		bBrowse.grid(row = 0, column = 1, padx=3, pady=3)
		for button in buttonInfo:
			# Set cancel button and location
			b = Button(self, text = button['text'], width = button['width'], command = button['function'])
			b.pack(side = RIGHT, padx=4, pady=2)

	# def getFilePath(self):
	# 	"""Get file Path from file path entry."""

	# 	self.filePath = askopenfilename(filetypes = (fileType, ("All files", "*.*")), parent = self.root)
		
	# 	fileType1 = ("Excel Workbook", "*.xlsx")
	# 	fileType2 = ("Excel Macro-Enabled Workbook", "*.xlsm")
	# 	self.filePath = askopenfilename(filetypes = (fileType2, fileType1, ("All files", "*.*")), parent = self.root)

	# 	# Once self.filePath gets a filepath, delete what's in the entry and put self.filePath into the entry
	# 	self.filePathEntry.delete(0, 'end')
	# 	self.filePathEntry.insert(0, self.filePath)

	def checkAndGetPath(self):
		"""A set of instructions are ran when ok button is clicked. 
	
		"""

		# Get a file path from browe entry
		self.filePath = self.filePathEntry.get()
		# When entry is empty						
		if self.filePath == "":
			self.emptyFileNameWarning()
			return False
		# When file path is invaild meaning not a legit file
		elif not exists(self.filePath):
			self.incorrectFileNameWarning()
			return False
		return True

	def closeWindow(self):
		"""Close the toplevel window."""

		self.root.destroy()

	def back(self):
		"""Go back to the first interface."""

		# Close the current window and unhide the first interface


	def incorrectFileNameWarning(self):
		"""Warning when file path is incorrect(file does not exist)."""

		messagebox.showinfo("Warning", "File does not exist!", parent = self.root)

	def emptyFileNameWarning(self):
		"""Warning when file path entry is empty but ok is clicked."""

		messagebox.showinfo("Warning", "No files selected!", parent = self.root)

	def popErrorMessage(self, message):
		"""Show error message in a pop up window.

			Parameters
			----------
			message: string
				An error message.
		"""

		messagebox.showinfo("Warning", message, parent = self.root)


class browseSheet(browse):

	def __init__(self, root):
		browse.__init__(self, root)
		self.nextFrame = None
		# b = browseFolder(root)
		# c = saveXML(root)

	def initGUI(self):
		super().initGUI()
		button = [{'text': 'Cancel', 'width': 10, 'function': self.closeWindow}, {'text': 'Next', 'width': 7, 'function': self.toBrowseFolder}]
		self.makeButtons(self.getFilePath, button)

	def getFilePath(self):
		fileType1 = ("Comma Delimited CSV", "*.CSV")
		fileType2 = ("Excel Workbook", "*.xlsx")
		fileType3 = ("Excel Macro-Enabled Workbook", "*.xlsm")
		fileType4 = ("All files", "*.*")
		
		self.filePath = askopenfilename(filetypes = (fileType1, fileType2, fileType3, fileType4), parent = self.root)

		# Once self.filePath gets a filepath, delete what's in the entry and put self.filePath into the entry
		self.filePathEntry.delete(0, 'end')
		self.filePathEntry.insert(0, self.filePath)

	def toBrowseFolder(self):
		if self.checkAndGetPath():
			self.lower()
			b = browseFolder(self.root)
			b.lift()

class browseFolder(browse):

	def __init__(self, root):
		browse.__init__(self, root)
		self.nextFrame = None

	def initGUI(self):
		super().initGUI()
		button = [{'text': 'Cancel', 'width': 10, 'function': self.closeWindow}, {'text': 'Next', 'width': 7, 'function': self.toSaveXML}, {'text': 'Back', 'width': 7, 'function': self.back}]
		self.makeButtons(self.getFilePath, button)

	def getFilePath(self):
		
		self.filePath = askdirectory(parent = self.root)
		print(self.filePath)

		# Once self.filePath gets a filepath, delete what's in the entry and put self.filePath into the entry
		self.filePathEntry.delete(0, 'end')
		self.filePathEntry.insert(0, self.filePath)

	def toSaveXML(self):
		self.checkAndGetPath()

class saveXML(browse):

	def __init__(self, root):
		browse.__init__(self, root)
		self.nextFrame = None

	def initGUI(self):
		super().initGUI()
		button = [{'text': 'Cancel', 'width': 10, 'function': self.closeWindow}, {'text': 'Save', 'width': 7, 'function': self.save}, {'text': 'Back', 'width': 7, 'function': self.back}]
		self.makeButtons(self.getFilePath, button)

	def getFilePath(self):
		fileType1 = ("XML Data", "*.xml")
		fileType2 = ("Text", "*.txt")
		fileType3 = ("All files", "*.*")

		self.filePath = asksaveasfilename(defaultextension= '.xml', filetypes = (fileType1, fileType2, fileType3), parent = self.root)
		print(self.filePath)

		# Once self.filePath gets a filepath, delete what's in the entry and put self.filePath into the entry
		self.filePathEntry.delete(0, 'end')
		self.filePathEntry.insert(0, self.filePath)

	def save(self):
		print(self.filePath)









window = Tk()
window.title("CAKE")
# Create a first object
# a = browseSheet(window)
# a.lift
# b = browseFolder(window)
c = saveXML(window)
# Launch
window.mainloop()
