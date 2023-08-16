import main
import idsettings as settings
import pandas as pd
from PIL import Image, ImageFont, ImageDraw
import openpyxl
from openpyxl_image_loader import SheetImageLoader


def compile_names(excel):
    df = pd.read_excel(excel, header=None)

    column = df[settings.NAME]
    if settings.MIDDLE_NAME != '':
        column2 = df[settings.MIDDLE_NAME]
    names = []
    i = 0
    for name in column:
        temp_name = ' '.join(elem.capitalize() for elem in name.split())
        if settings.MIDDLE_NAME != '':
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
    qrcodes = qr.generate(invert=settings.INVERT, split=settings.SPLIT, save=False, clear=settings.DELETE_PREV, border=settings.BORDER_SIZE)
    return qrcodes


def generate_name(size, message, font, fontColor):
    W, H = size
    image = Image.new('RGBA', size, (255, 255, 255,0))
    draw = ImageDraw.Draw(image)
    _, _, w, h = draw.textbbox((0, 0), message, font=font)
    draw.text(((W-w)/2, (H-h)/2), message, font=font, fill=fontColor)
    return image


def generate_id(template, name, picture, qrcode):
    id = Image.open(template)
    name = name.split(', ')
    formatted_name = name[1] + ' ' + name[0]
    id.paste(generate_name(settings.NAME_SIZE, formatted_name, ImageFont.truetype(settings.FONT, settings.FONT_SIZE), settings.FONT_COLOR), settings.NAME_POS)
    
    pic_pos = (settings.PICTURE_POS[0], settings.PICTURE_POS[1])
    pic_size = (settings.PICTURE_POS[2], settings.PICTURE_POS[3])
    if picture != None:
        id.paste(picture.resize(pic_size), pic_pos)
    
    qr_pos = (settings.QR_POS[0], settings.QR_POS[1])
    qr_size = (settings.QR_POS[2], settings.QR_POS[3])
    if qrcode != None:
        id.paste(qrcode.resize(qr_size), qr_pos)

    id.save('exports/test.png')


def generate_ids():
    names = compile_names(settings.EXCEL)
    str_names = "\n".join(names)
    qrcodes = generate_qr(str_names)

    i = 1
    image = load_image(settings.EXCEL, i)
    generate_id(settings.TEMPLATE, names[i].upper(), image, qrcodes[i])


    # image.save("exports/pic1.png")
    # print(len(qrcodes))
    # print(qrcodes)
    # for code in qrcodes:
    #     code.save("exports/" + names[qrcodes.index(code)] + ".png")
    #     print(code)

generate_ids()