import main
import idsettings as settings
import pandas as pd
from PIL import Image
import openpyxl
from openpyxl_image_loader import SheetImageLoader

name = """test1
test2"""


def compile_names(excel):
    df = pd.read_excel(excel, header=None)

    column = df[settings.NAME]
    column2 = df[settings.MIDDLE_NAME]
    names = []
    i = 0
    for name in column:
        temp_name = ' '.join(elem.capitalize() for elem in name.split())
        if column2[i] != '':
            temp_name = temp_name + " " + column2[i].upper() + "."
        names.append(temp_name)
        i += 1
    return names


def load_image(excel, index):
    
    wb = openpyxl.load_workbook(excel)
    sheet = wb.worksheets[0]
    image_loader = SheetImageLoader(sheet)
    image = None
    if image_loader.image_in(settings.PICTURE+str(index+1)):
        image = image_loader.get(settings.PICTURE+str(index+1))
    return image

def generate_qr(names):
    qr = main.SimpleQR(names, settings.LINK)
    qrcodes = qr.generate(invert=settings.INVERT, split=settings.SPLIT, save=False)
    return qrcodes


def generate_ids():
    names = compile_names(settings.EXCEL)
    str_names = "\n".join(names)
    qrcodes = generate_qr(str_names)
    image = load_image(settings.EXCEL, 1)

    
    # image.save("exports/pic1.png")
    # print(len(qrcodes))
    # print(qrcodes)
    # for code in qrcodes:
    #     code.save("exports/" + names[qrcodes.index(code)] + ".png")
    #     print(code)

generate_ids()