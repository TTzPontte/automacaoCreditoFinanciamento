import os
import glob
import os.path
from PIL import Image

def checkFolder(path):
    folder_path = path

    image_list = glob.glob(folder_path + '/*.jpg') + glob.glob(folder_path + '/*.jpeg') + glob.glob(folder_path + '/*.png')

    if len(image_list) > 0:
        for image_file in image_list:
            if os.path.isfile(image_file):
                image = Image.open(image_file)
                image = image.convert('RGB',colors=255)
                pdf_file = os.path.splitext(image_file)[0] + ".pdf"
                image.save(pdf_file,"PDF" )
                os.remove(image_file)