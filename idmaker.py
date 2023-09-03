import main
import pandas as pd
import openpyxl
import os
import webbrowser
from configparser import ConfigParser
from PIL import Image, ImageFont, ImageDraw, ImageOps
from openpyxl.utils import column_index_from_string
from openpyxl_image_loader import SheetImageLoader
from openpyxl import Workbook


class Settings(object):
    pass


class IDMaker():
    def __init__(self, settings):
        self.cfg = ConfigParser()
        self.cfg.read(settings)
        if not self.cfg.has_section('simpleqr'):
            raise Exception("\"simpleqr\" section is missing in config file")
        if not self.cfg.has_section('excel'):
            raise Exception("\"excel\" section is missing in config file")
        if not self.cfg.has_section('id'):
            raise Exception("\"id\" section is missing in config file")
            
        self.settings = Settings()
        
        self.settings.SPLIT = (self.cfg.get('simpleqr','SPLIT') == "True") if self.cfg.has_option('simpleqr', 'SPLIT') else None
        self.settings.INVERT = (self.cfg.get('simpleqr','INVERT') == "True") if self.cfg.has_option('simpleqr', 'INVERT') else None
        self.settings.DELETE_PREV = (self.cfg.get('simpleqr','DELETE_PREV') == "True") if self.cfg.has_option('simpleqr', 'DELETE_PREV') else None
        self.settings.CLEAR_EXCEL = (self.cfg.get('simpleqr','CLEAR_EXCEL') == "True") if self.cfg.has_option('simpleqr', 'CLEAR_EXCEL') else None
        self.settings.BORDER_SIZE = self.cfg.getint('simpleqr','BORDER_SIZE') if self.cfg.has_option('simpleqr', 'BORDER_SIZE') else None
        self.settings.LINK = self.cfg.get('simpleqr','LINK') if self.cfg.has_option('simpleqr', 'LINK') else None
        
        self.settings.EXCEL = self.cfg.get('excel','EXCEL') if self.cfg.has_option('excel', 'EXCEL') else None
        self.settings.RANGE = self.cfg.getint('excel','RANGE') if self.cfg.has_option('excel', 'RANGE') else None
        self.settings.NAME = self.cfg.get('excel','NAME') if self.cfg.has_option('excel', 'NAME') else None
        self.settings.MIDDLE_NAME = self.cfg.get('excel','MIDDLE_NAME') if self.cfg.has_option('excel', 'MIDDLE_NAME') else None
        self.settings.PICTURE = self.cfg.get('excel','PICTURE') if self.cfg.has_option('excel', 'PICTURE') else None
        self.settings.ROOM_NUMBER = self.cfg.get('excel','ROOM_NUMBER') if self.cfg.has_option('excel', 'ROOM_NUMBER') else None
       
        self.settings.TEMPLATE = self.cfg.get('id','TEMPLATE') if self.cfg.has_option('id', 'TEMPLATE') else None
        self.settings.FONT = self.cfg.get('id','FONT') if self.cfg.has_option('id', 'FONT') else None
        self.settings.NAME_SIZE = self.cfg.getint('id','NAME_SIZE') if self.cfg.has_option('id', 'NAME_SIZE') else None
        self.settings.NAME_COLOR = tuple(map(int, self.cfg.get('id','NAME_COLOR').replace(' ', '').split(','))) if self.cfg.has_option('id', 'NAME_COLOR') else None
        self.settings.PICTURE_POS = tuple(map(int, self.cfg.get('id','PICTURE_POS').replace(' ', '').split(','))) if self.cfg.has_option('id', 'PICTURE_POS') else None
        self.settings.ASPECT_RATIO = self.cfg.get('id','ASPECT_RATIO') if self.cfg.has_option('id', 'ASPECT_RATIO') else None
        self.settings.MISSING_PICTURE = self.cfg.get('id','MISSING_PICTURE') if self.cfg.has_option('id', 'MISSING_PICTURE') else None
        self.settings.QR_POS = tuple(map(int, self.cfg.get('id','QR_POS').replace(' ', '').split(','))) if self.cfg.has_option('id', 'QR_POS') else None
        self.settings.NAME_POS = tuple(map(int, self.cfg.get('id','NAME_POS').replace(' ', '').split(','))) if self.cfg.has_option('id', 'NAME_POS') else None
        self.settings.RN_SIZE = self.cfg.getint('id','RN_SIZE') if self.cfg.has_option('id', 'RN_SIZE') else None
        self.settings.RN_COLOR = self.cfg.get('id','RN_COLOR') if self.cfg.has_option('id', 'RN_COLOR') else None
        self.settings.RN_POS = tuple(map(int, self.cfg.get('id','RN_POS').replace(' ', '').split(','))) if self.cfg.has_option('id', 'RN_POS') else None
        
        for var in vars(self.settings):
            if getattr(self.settings, var) == None:
                raise Exception(var+" is missing from config file")


    def compile_names(self, excel):
        df = pd.read_excel(excel, header=None)

        column = df[column_index_from_string(self.settings.NAME)-1]
        if self.settings.MIDDLE_NAME != '':
            column2 = df[self.settings.MIDDLE_NAME]
        names = []
        i = 0
        for name in column:
            temp_name = ' '.join(elem.capitalize() for elem in name.split())
            if self.settings.MIDDLE_NAME != '':
                if column2[i] != '':
                    temp_name = temp_name + " " + column2[i].upper() + "."
            names.append(temp_name)
            i += 1
        return names


    def compile_rooms(self, excel):
        df = pd.read_excel(excel, header=None)
        rooms = df[column_index_from_string(self.settings.ROOM_NUMBER)-1].to_list()
        return rooms


    def load_images(self, excel, column, total):
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

    def generate_qr(self, names):
        qr = main.SimpleQR(names, self.settings.LINK)
        qrcodes = qr.generate(invert=self.settings.INVERT, split=self.settings.SPLIT, save=False, clear=self.settings.DELETE_PREV, border=self.settings.BORDER_SIZE)
        return qrcodes


    def generate_text(self, size, message, font, fontColor):
        W, H = size
        image = Image.new('RGBA', size, (255, 255, 255,0))
        draw = ImageDraw.Draw(image)
        _, _, w, h = draw.textbbox((0, 0), str(message), font=font)
        draw.text(((W-w)/2, (H-h)/2), str(message), font=font, fill=fontColor)
        return image


    def generate_id(self, template, name, picture, qrcode, room):
        id = Image.open(template)

        name_split = name.split(', ')
        formatted_name = name_split[1] + ' ' + name_split[0]
        name_pos = (self.settings.NAME_POS[0], self.settings.NAME_POS[1])
        name_size = (self.settings.NAME_POS[2], self.settings.NAME_POS[3])
        id.paste(self.generate_text(name_size, formatted_name, ImageFont.truetype(self.settings.FONT, self.settings.NAME_SIZE), self.settings.NAME_COLOR), name_pos)
        
        if room != '':
            room_pos = (self.settings.RN_POS[0], self.settings.RN_POS[1])
            room_size = (self.settings.RN_POS[2], self.settings.RN_POS[3])
            id.paste(self.generate_text(room_size, room, ImageFont.truetype(self.settings.FONT, self.settings.RN_SIZE), self.settings.RN_COLOR), room_pos)

        if picture != None:
            pic_pos = (self.settings.PICTURE_POS[0], self.settings.PICTURE_POS[1])
            pic_size = (self.settings.PICTURE_POS[2], self.settings.PICTURE_POS[3])
            if self.settings.ASPECT_RATIO.lower() == 'stretch':
                id.paste(picture.resize(pic_size), pic_pos)
            else: #keep
                resized_pic = ImageOps.contain(picture, pic_size)
                new_pos = (int((pic_size[0]/2)-(resized_pic.size[0]/2)+pic_pos[0]),
                        int((pic_size[1]/2)-(resized_pic.size[1]/2)+pic_pos[1]))
                id.paste(resized_pic, new_pos)
        elif self.settings.MISSING_PICTURE.lower() == 'skip':
            return

        if qrcode != None:
            qr_pos = (self.settings.QR_POS[0], self.settings.QR_POS[1])
            qr_size = (self.settings.QR_POS[2], self.settings.QR_POS[3])
            id.paste(qrcode.resize(qr_size), qr_pos)

        return id


    def clear_excel(self):
        Wb = Workbook()
        sheet = Wb.worksheets[0]
        sheet.column_dimensions[self.settings.NAME].width = 25
        sheet.column_dimensions[self.settings.PICTURE].width = 50
        
        for i in range(200):
            sheet.row_dimensions[i+1].height = 200

        Wb.save(self.settings.EXCEL)


    def generate_ids(self):
        if not os.path.isfile(self.settings.EXCEL):
            self.clear_excel()
            print("Excel File not found. Generating a new excel file.")
            return

        print("Extracting Names")
        names = self.compile_names(self.settings.EXCEL)
        if self.settings.RANGE > 0:
            names = names[:self.settings.RANGE]
        str_names = "\n".join(names)

        print("Generating QR Codes")
        qrcodes = self.generate_qr(str_names)
        
        pictures = self.load_images(self.settings.EXCEL, self.settings.PICTURE, len(names))

        rooms = []
        if self.settings.ROOM_NUMBER != '':
            rooms = self.compile_rooms(self.settings.EXCEL)

        print("Generating IDs")
        skipped = []
        for i in range(len(names)):
            print("Generating IDs (" + str(i+1) + "/" + str(len(names)) + ")")

            room = ''
            if rooms != []:
                room = rooms[i]

            id = self.generate_id(self.settings.TEMPLATE, names[i].upper(), pictures[i], qrcodes[i], room)
            if id == None:
                skipped.append(names[i])
            else:
                id.save('exports/'+names[i]+'.png')
        
        print("Missing Pictures:\n" + '\n'.join(skipped) + '\n')
        print("Done! (" + str(len(names)-len(skipped)) + "/" + str(len(names)) + ") have been generated successfully.")
        webbrowser.open('file://' + os.getcwd().replace('\\', '/') + '/exports')
        
        if self.settings.CLEAR_EXCEL:
            self.clear_excel()