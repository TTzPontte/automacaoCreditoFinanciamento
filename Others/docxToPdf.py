from docx2pdf import convert
from PyPDF2 import PdfFileReader, PdfFileMerger
import os

def convertDocxToPDF(path): 
    docx_files = [f for f in os.listdir(path) if f.endswith("docx")]
    i = 0
    
    for file in docx_files:
        file = os.path.join(path, file)
        i = i + 1
        convert(file, os.path.join(path, f"PDF{1}.pdf"))
    
    mergePDF(path)

def mergePDF(pathImovel):
    files_dir = pathImovel
    pdf_files = [f for f in os.listdir(files_dir) if f.endswith("pdf")]
    merger = PdfFileMerger()

    for filename in pdf_files:
        merger.append(PdfFileReader(os.path.join(files_dir, filename), "rb"))
        pathFile = pathImovel+ f"\{filename}"
        os.remove(pathFile)

    merger.write(os.path.join(files_dir, "exportMergePDF.pdf"))


folderPath = r'pathAqui'

convertDocxToPDF(folderPath)
