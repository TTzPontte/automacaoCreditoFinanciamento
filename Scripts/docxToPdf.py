from docx2pdf import convert
from PyPDF2 import PdfFileReader, PdfFileMerger
import os

def convertDocxToPDF(path): 
    docx_files = [f for f in os.listdir(path) if f.endswith("docx")]
    i = 0
    
    for file in docx_files:
        file = os.path.join(path, file)
        i = i + 1
        convert(file, rf'C:\Users\Matheus\Desktop\Finanças P3\teste{i}.pdf')
    
    mergePDF(path)

def mergePDF(pathImovel):
    files_dir = pathImovel
    pdf_files = [f for f in os.listdir(files_dir) if f.endswith("pdf")]
    merger = PdfFileMerger()

    for filename in pdf_files:
        merger.append(PdfFileReader(os.path.join(files_dir, filename), "rb"))
        pathFile = pathImovel+ f"\{filename}"
        os.remove(pathFile)

    merger.write(os.path.join(files_dir, "Colas Mágicas.pdf"))


folderPath = r'C:\Users\Matheus\Desktop\Finanças P3'

convertDocxToPDF(folderPath)
