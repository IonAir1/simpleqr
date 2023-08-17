import main
import idsettings as settings
import pandas as pd
from PIL import Image, ImageFont, ImageDraw, ImageOps
import openpyxl
from openpyxl.utils import column_index_from_string
from openpyxl_image_loader import SheetImageLoader


def compile_names(excel):
    df = pd.read_excel(excel, header=None)

    column = df[column_index_from_string(settings.NAME)-1]
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
    name_pos = (settings.NAME_POS[0], settings.NAME_POS[1])
    name_size = (settings.NAME_POS[2], settings.NAME_POS[3])
    id.paste(generate_name(name_size, formatted_name, ImageFont.truetype(settings.FONT, settings.FONT_SIZE), settings.FONT_COLOR), name_pos)
    
    if picture != None:
        pic_pos = (settings.PICTURE_POS[0], settings.PICTURE_POS[1])
        pic_size = (settings.PICTURE_POS[2], settings.PICTURE_POS[3])
        if settings.ASPECT_RATIO.lower() == 'stretch':
            id.paste(picture.resize(pic_size), pic_pos)
        else: #keep
            resized_pic = ImageOps.contain(picture, pic_size)
            new_pos = (int((pic_size[0]/2)-(resized_pic.size[0]/2)+pic_pos[0]),
                       int((pic_size[1]/2)-(resized_pic.size[1]/2)+pic_pos[1]))
            id.paste(resized_pic, new_pos)

    if qrcode != None:
        qr_pos = (settings.QR_POS[0], settings.QR_POS[1])
        qr_size = (settings.QR_POS[2], settings.QR_POS[3])
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