from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import os
from PyPDF2 import PdfFileReader
from pdf2docx import Converter

class Window(Frame):

	def __init__(self, master=None):
		self.targetPdf = ""
		self.is3Page = IntVar()
		Frame.__init__(self, master)
		self.master = master
		self.init_window()

	def init_window(self):
		# Defines the positions of all the elements in the window

		self.master.title("Extract back 2 pages from PDF")
		self.pack(fill=BOTH, expand=1)
		
		selectInputButton = Button(self, text="Select input folder", command=self.selectInputFolder)
		is3PageButton = Checkbutton(self, text = "Joint case?", variable=self.is3Page, offvalue=1, onvalue=3)
		startButton = Button(self, text="Start", command=self.startFunction)
		quitButton = Button(self, text="Quit", command=root.destroy)
		
		selectInputButton.pack(padx=7, pady=7)
		is3PageButton.pack(padx=7, pady=7)
		startButton.pack(padx=7, pady=7)
		quitButton.pack(padx=7, pady=20)

		# I've changed the checkbox's values to 1 and 3 (tkinter's default is 0 and 1), but it'll remain at 0 until it's first interacted with.
		# Deselecting it here forces the value to be 1, preventing it from becoming an issue.
		is3PageButton.deselect()

	def selectInputFolder(self):
		# Sets the directory to run on
		self.targetFolder = filedialog.askdirectory(initialdir = os.path.dirname(os.path.realpath(__file__)), title = "Select folder")

	def populateFileList_recursive(self):
		# Walk through the selected input folder and all its subfolders, and return a list of filepaths matching the given conditions to run the extractor on.
		folderStructure = [os.path.join(dp, f) for dp, dn, fn in os.walk(self.targetFolder) for f in fn]

		foundPdfs = []

		for file in folderStructure:
			# Script only needs to be run on specific files - PDFs, where the filename contains the word "welcome pack".
			# Only add files matching these conditions to the list of items to run on.
			fName = os.path.basename(file)
			fExt = os.path.splitext(fName)
			if fExt[1] == ".pdf" and "welcome pack" in fExt[0].lower():
				foundPdfs.append(file)

		return foundPdfs

	def startFunction(self):
		# Start the main functionality. Populate the list of PDF files to run on, then for each read it and output its back page as a .docx
		# Requires self.targetFolder to be set (via selectInputFolder from the main menu) so it can find the input

		if not self.targetFolder:
			messagebox.showinfo("Error", "No input folder given")
			return

		pdfsList = self.populateFileList_recursive()

		for inputPdfPath in pdfsList:
			# Generate the filepath for the output
			inputDirectory = os.path.dirname(inputPdfPath)
			outputPath = "ExtractedLetter_%s" %(os.path.basename(inputPdfPath))
			outputPath =  "%s.docx" %(os.path.splitext(outputPath)[0])
			outputPath = os.path.join(inputDirectory, outputPath)

			# Read the PDF file and get the number of pages it contains.
			pdfReader = PdfFileReader(inputPdfPath)
			maxPage = pdfReader.getNumPages()
			
			# The number of pages before the last page to start the extraction. Use 3 instead of 1 for some cases
			# Handled via making 1 and 3 the on and off value of the checkbox, and directly extracting the value here.
			backPageOffset = self.is3Page.get()

			# Convert the input PDF to a .docx file.
			# Start the conversion at a given point (either the last page or the last 3 pages, decided by backPageOffset), and end it at the end of the document.
			cv = Converter(inputPdfPath)
			cv.convert(outputPath, start=maxPage - backPageOffset, end=None)
			cv.close()

			messagebox.showinfo("Done", "Created file %s" %(outputPath))

root = Tk()
root.geometry("200x200")
app = Window(root)

root.mainloop()
