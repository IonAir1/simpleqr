import main
import idsettings as settings
import pandas as pd
import openpyxl
import os
import webbrowser
from PIL import Image, ImageFont, ImageDraw, ImageOps
from openpyxl.utils import column_index_from_string
from openpyxl_image_loader import SheetImageLoader
from openpyxl import Workbook


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


def compile_rooms(excel):
    df = pd.read_excel(excel, header=None)
    rooms = df[column_index_from_string(settings.ROOM_NUMBER)-1].to_list()
    return rooms


def load_images(excel, column, total):
    wb = openpyxl.load_workbook(excel)
    sheet = wb.worksheets[0]
    image_loader = SheetImageLoader(sheet)
    images = []
    for i in range(total):
        print("Extracting images from excel file (" + str(i+1) + "/" + str(total) + ")")
        if image_loader.image_in(str(column)+str(i+1)):
            images.append(image_loader.get(column+str(i+1)))
        else:
            images.append(None)
    return images

def generate_qr(names):
    qr = main.SimpleQR(names, settings.LINK)
    qrcodes = qr.generate(invert=settings.INVERT, split=settings.SPLIT, save=False, clear=settings.DELETE_PREV, border=settings.BORDER_SIZE)
    return qrcodes


def generate_text(size, message, font, fontColor):
    W, H = size
    image = Image.new('RGBA', size, (255, 255, 255,0))
    draw = ImageDraw.Draw(image)
    _, _, w, h = draw.textbbox((0, 0), message, font=font)
    draw.text(((W-w)/2, (H-h)/2), message, font=font, fill=fontColor)
    return image


def generate_id(template, name, picture, qrcode, room):
    id = Image.open(template)

    name_split = name.split(', ')
    formatted_name = name_split[1] + ' ' + name_split[0]
    name_pos = (settings.NAME_POS[0], settings.NAME_POS[1])
    name_size = (settings.NAME_POS[2], settings.NAME_POS[3])
    id.paste(generate_text(name_size, formatted_name, ImageFont.truetype(settings.FONT, settings.NAME_SIZE), settings.NAME_COLOR), name_pos)
    
    if room != '':
        room_pos = (settings.RN_POS[0], settings.RN_POS[1])
        room_size = (settings.RN_POS[2], settings.RN_POS[3])
        id.paste(generate_text(room_size, room, ImageFont.truetype(settings.FONT, settings.RN_SIZE), settings.RN_COLOR), room_pos)

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
    elif settings.MISSING_PICTURE.lower() == 'skip':
        return

    if qrcode != None:
        qr_pos = (settings.QR_POS[0], settings.QR_POS[1])
        qr_size = (settings.QR_POS[2], settings.QR_POS[3])
        id.paste(qrcode.resize(qr_size), qr_pos)

    return id


def clear_excel():
    Wb = Workbook()
    sheet = Wb.worksheets[0]
    sheet.column_dimensions[settings.NAME].width = 25
    sheet.column_dimensions[settings.PICTURE].width = 50
    
    for i in range(200):
        sheet.row_dimensions[i+1].height = 200

    Wb.save(settings.EXCEL)


def generate_ids():
    if not os.path.isfile(settings.EXCEL):
        clear_excel()
        print("Excel File not found. Generating a new excel file.")
        return

    print("Extracting Names")
    names = compile_names(settings.EXCEL)
    if settings.RANGE > 0:
        names = names[:settings.RANGE]
    str_names = "\n".join(names)

    print("Generating QR Codes")
    qrcodes = generate_qr(str_names)
    
    pictures = load_images(settings.EXCEL, settings.PICTURE, len(names))

    rooms = []
    if settings.ROOM_NUMBER != '':
        rooms = compile_rooms(settings.EXCEL)

    print("Generating IDs")
    skipped = []
    for i in range(len(names)):
        print("Generating IDs (" + str(i+1) + "/" + str(len(names)) + ")")

        room = ''
        if rooms != []:
            room = rooms[i]

        id = generate_id(settings.TEMPLATE, names[i].upper(), pictures[i], qrcodes[i], room)
        if id == None:
            skipped.append(names[i])
        else:
            id.save('exports/'+names[i]+'.png')
    
    print("Missing Pictures:\n" + '\n'.join(skipped) + '\n')
    print("Done! (" + str(len(names)-len(skipped)) + "/" + str(len(names)) + ") have been generated successfully.")
    webbrowser.open('file://' + os.getcwd().replace('\\', '/') + '/exports')
    
    if settings.CLEAR_EXCEL:
        clear_excel()


# clear_excel()
generate_ids()