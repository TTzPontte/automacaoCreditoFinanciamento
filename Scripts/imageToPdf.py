from PIL import Image
from PyPDF2 import PdfFileReader, PdfFileMerger
import warnings
warnings.filterwarnings("ignore", category=DeprecationWarning)
import os 
list = {}

def transformImageToPDF(path):
    image_1 = Image.open(path)
    im_1 = image_1.convert('RGB')
    saida = path
    saida = saida.replace('.png','.pdf')
    saida = saida.replace('.jpeg','.pdf')
    saida = saida.replace('.jpg','.pdf')
    
    im_1.save(saida)

    os.remove(path)

def mergePDF(pathImovel):
    files_dir = pathImovel
    pdf_files = [f for f in os.listdir(files_dir) if f.endswith("pdf")]
    merger = PdfFileMerger()

    for filename in pdf_files:
        merger.append(PdfFileReader(os.path.join(files_dir, filename), "rb"))
        pathFile = pathImovel+ f"\{filename}"
        os.remove(pathFile)

    merger.write(os.path.join(files_dir, "Fotos do Imovel.pdf"))
