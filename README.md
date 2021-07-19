# PDF page extractor
 Walks through a given folder and its subfolders and extracts the back pages of PDFs matching a certain naming scheme as .docx files.
 Output .docx files are placed in the same folder as their original .pdf. 

# Usage
 Select the folder to walk through with the "Select input folder" button.
 If the "Joint case?" tickbox is checked, the back 3 pages are extracted instead of just the final page of the document.
 Naming scheme for the output is "ExtractedLetter_[original filname].docx"

 The use case for this script has other, non-desired PDFs in the same folder structure. The files we want this to run on are identified by a naming scheme:
 Currently the naming scheme used for the input pdfs is "Welcome pack" (not case-sensitive) in the filename.