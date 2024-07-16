from configparser import ConfigParser
from openpyxl.utils import column_index_from_string
from openpyxl_image_loader import SheetImageLoader
from openpyxl import Workbook
from PIL import Image, ImageFont, ImageDraw, ImageOps
import openpyxl
import os
import pandas as pd
import qrcode
import shutil
import webbrowser


class Config():
    def __init__(self, file):
        self.file = file
    
    cfg = ConfigParser()
    text = ''
    url = ''
    invert = False
    split = False
    
    def save(self, key, value):
        self.cfg.read(self.file)
        if not self.cfg.has_section('main'):
            self.cfg.add_section('main')
        self.cfg.set('main', str(key), str(value))
        with open(self.file, 'w') as f: #save
            self.cfg.write(f)
        print(key,"  ", value)

    def load(self):
        self.cfg.read(self.file)
        if not self.cfg.has_section('main'):
            self.cfg.add_section('main')
        if self.cfg.has_option('main', 'text'):
            self.text = self.cfg.get('main','text')
        if self.cfg.has_option('main', 'url'):
            self.url = self.cfg.get('main','url')
        if self.cfg.has_option('main', 'invert'):
            self.invert = (self.cfg.get('main','invert') == "True")
        if self.cfg.has_option('main', 'split'):
            self.split = (self.cfg.get('main','split') == "True")

        
class SimpleQR():
    def __init__(self, text, url, **kwargs):
        
        self.text = os.linesep.join([s for s in text.splitlines() if s])
        self.url = url
        self.progress = 0
        self.output_path = kwargs.get('path','exports/')


    def generate(self, **kwargs):
        self.progress = 0
        #parse filenames
        filenames = self.text.replace('\\r','').replace('\r','').split('\n')
        
        #generate prefill links
        urls = self.generate_links(filenames, self.url, **kwargs)
        
        #clear output folder
        if not os.path.exists(self.output_path):
            os.makedirs(self.output_path)
        if kwargs.get('clear', True):
            for f in os.listdir(self.output_path):
                if not f.endswith(".png"):
                    continue
                os.remove(os.path.join(self.output_path, f))
        
        #generate qr codes
        a = (1/len(filenames))*100
        qrcodes = []
        for n in range(len(filenames)):
            qrcodes.append(self.generate_qr(self.output_path, urls[n], filenames[n], kwargs.get('invert', False), kwargs.get('box', 10), kwargs.get('border', 4), kwargs.get('save', True)))
            self.progress += a

        self.progress = 100
        return qrcodes
    
    
    #generate links
    def generate_links(self, names, url, **kwargs):
        if kwargs.get('replace', True):
            urls = []
            rooms = kwargs.get('room', None)

            for i in range(len(names)):
                output_url = ""
                if kwargs.get('split', False) in ['1',True]: #check if split name is on
                    last_first = names[i].split(',')
                    new_url = url.replace('=lastname','='+last_first[0].strip().replace(' ', '%20'))

                    if len(last_first) > 1: #check if first name exists
                        if last_first[1].strip()[-1] == '.':
                            first_mid = last_first[1].strip().split(' ')
                            final_url = new_url
                            for name in first_mid: #does the first name string contain a middle name
                                if name[-1] == '.':
                                    final_url = final_url.replace('addmiddlename','%20'+name.strip().replace(' ', '%20')+'addmiddlename')
                                    final_url = final_url.replace('=middlename','='+name.strip().replace(' ', '%20')+'addmiddlename')
                                else:
                                    final_url = final_url.replace('addfirstname','%20'+name.strip().replace(' ', '%20')+'addfirstname')
                                    final_url = final_url.replace('=firstname','='+name.strip().replace(' ', '%20')+'addfirstname')
                            output_url = final_url.replace('addfirstname','').replace('addmiddlename','').replace('firstname','').replace('middlename','')

                        else:
                            new_url = new_url.replace('=firstname','='+last_first[1].strip().replace(' ', '%20'))
                            output_url = new_url.replace('firstname','').replace('middlename','')
                    else:
                        output_url = new_url.replace('firstname','').replace('middlename','')

                else:
                    output_url = url.replace('=name','='+name.replace(' ', '%20'))
                
                if rooms != None: #add rooms if required
                    output_url = output_url.replace('=room', '='+str(rooms[i]))

                urls.append(output_url)

            return urls
        elif not(kwargs.get('replace', False)):
            return names


    #generates an image of a qr code that links to specified url
    def generate_qr(self, output_path, url, name, invert, box, border, save):
        qr = qrcode.QRCode( #qr code properties
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=box,
        border=border,
        )

        qr.add_data(url) #qr code data
        qr.make(fit=True)
        if invert:
            img = qr.make_image(fill_color="black", back_color="white").convert('RGB')
        else:
            img = qr.make_image(fill_color="white", back_color="black").convert('RGB')
        if save:
            img.save(output_path.replace('\\','/') + str(name).replace('/','').replace(':','').replace('.', '') + ".png") #save
        return img


class IDMaker():
    def __init__(self, settings):
        #read config file
        if not os.path.isfile(settings):
            raise Exception("Config File \""+settings+"\" is not found")

        cfg = ConfigParser()
        cfg.read(settings)
        if not cfg.has_section('simpleqr'):
            raise Exception("\"simpleqr\" section is missing in config file")
        if not cfg.has_section('excel'):
            raise Exception("\"excel\" section is missing in config file")
        if not cfg.has_section('id'):
            raise Exception("\"id\" section is missing in config file")
        
        self.SPLIT = (cfg.get('simpleqr','SPLIT') == "True") if cfg.has_option('simpleqr', 'SPLIT') else None
        self.INVERT = (cfg.get('simpleqr','INVERT') == "True") if cfg.has_option('simpleqr', 'INVERT') else None
        self.DELETE_PREV = (cfg.get('simpleqr','DELETE_PREV') == "True") if cfg.has_option('simpleqr', 'DELETE_PREV') else None
        self.CLEAR_EXCEL = (cfg.get('simpleqr','CLEAR_EXCEL') == "True") if cfg.has_option('simpleqr', 'CLEAR_EXCEL') else None
        self.CONFIRM = (cfg.get('simpleqr','CONFIRM') == "True") if cfg.has_option('simpleqr', 'CONFIRM') else None
        self.BORDER_SIZE = cfg.getint('simpleqr','BORDER_SIZE') if cfg.has_option('simpleqr', 'BORDER_SIZE') else None
        self.LINK = cfg.get('simpleqr','LINK') if cfg.has_option('simpleqr', 'LINK') else None
        
        self.EXCEL = cfg.get('excel','EXCEL') if cfg.has_option('excel', 'EXCEL') else None
        self.START_ROW = cfg.getint('excel','START_ROW') if cfg.has_option('excel', 'START_ROW') else None
        self.END_ROW = cfg.getint('excel','END_ROW') if cfg.has_option('excel', 'END_ROW') else None
        self.NAME = cfg.get('excel','NAME') if cfg.has_option('excel', 'NAME') else None
        self.MIDDLE_NAME = cfg.get('excel','MIDDLE_NAME') if cfg.has_option('excel', 'MIDDLE_NAME') else None
        self.PICTURE = cfg.get('excel','PICTURE') if cfg.has_option('excel', 'PICTURE') else None
        self.ROOM_NUMBER = cfg.get('excel','ROOM_NUMBER') if cfg.has_option('excel', 'ROOM_NUMBER') else None
       
        self.TEMPLATE = cfg.get('id','TEMPLATE') if cfg.has_option('id', 'TEMPLATE') else None
        self.FONT = cfg.get('id','FONT') if cfg.has_option('id', 'FONT') else None
        self.NAME_SIZE = cfg.getint('id','NAME_SIZE') if cfg.has_option('id', 'NAME_SIZE') else None
        self.NAME_COLOR = tuple(map(int, cfg.get('id','NAME_COLOR').replace(' ', '').split(','))) if cfg.has_option('id', 'NAME_COLOR') else None
        self.PICTURE_POS = tuple(map(int, cfg.get('id','PICTURE_POS').replace(' ', '').split(','))) if cfg.has_option('id', 'PICTURE_POS') else None
        self.ASPECT_RATIO = cfg.get('id','ASPECT_RATIO') if cfg.has_option('id', 'ASPECT_RATIO') else None
        self.MISSING_PICTURE = cfg.get('id','MISSING_PICTURE') if cfg.has_option('id', 'MISSING_PICTURE') else None
        self.QR_POS = tuple(map(int, cfg.get('id','QR_POS').replace(' ', '').split(','))) if cfg.has_option('id', 'QR_POS') else None
        self.NAME_POS = tuple(map(int, cfg.get('id','NAME_POS').replace(' ', '').split(','))) if cfg.has_option('id', 'NAME_POS') else None
        self.RN_SIZE = cfg.getint('id','RN_SIZE') if cfg.has_option('id', 'RN_SIZE') else None
        self.RN_COLOR = tuple(map(int, cfg.get('id','RN_COLOR').replace(' ', '').split(','))) if cfg.has_option('id', 'RN_COLOR') else None
        self.RN_POS = tuple(map(int, cfg.get('id','RN_POS').replace(' ', '').split(','))) if cfg.has_option('id', 'RN_POS') else None
        
        for var in vars(self):
            if getattr(self, var) == None:
                raise Exception(var+" is missing from config file")


    #collects the names from excel file to list
    def compile_names(self, excel):
        df = pd.read_excel(excel, header=None)

        column = df[column_index_from_string(self.NAME)-1]
        if self.MIDDLE_NAME != '':
            column2 = df[self.MIDDLE_NAME]
        names = []
        i = 0
        for name in column:
            temp_name = ' '.join(elem.capitalize() for elem in name.split())
            if self.MIDDLE_NAME != '':
                if column2[i] != '':
                    temp_name = temp_name + " " + column2[i].upper() + "."
            names.append(temp_name)
            i += 1
        return names


    #collects the room numbers from excel file to list
    def compile_rooms(self, excel):
        df = pd.read_excel(excel, header=None)
        rooms = df[column_index_from_string(self.ROOM_NUMBER)-1].to_list()
        return rooms


    #collects the pictures from excel file to list
    def load_images(self, excel, column, total):
        wb = openpyxl.load_workbook(excel)
        sheet = wb.worksheets[0]
        image_loader = SheetImageLoader(sheet)
        images = []
        for i in range(total):
            print("Extracting images from excel file ({}/{})".format(i+1, total))
            if image_loader.image_in(str(column)+str(i+1)):
                images.append(image_loader.get(column+str(i+1)))
            else:
                images.append(None)
        return images


    #generates text image to paste on id
    def generate_text(self, size, message, font, fontColor):
        W, H = size
        image = Image.new('RGB', size, (255, 255, 255))
        draw = ImageDraw.Draw(image)
        _, _, w, h = draw.textbbox((0, 0), str(message), font=font)
        draw.text(((W-w)/2, (H-h)/2), str(message), font=font, fill=fontColor)
        return image


    #paste id elements to template
    def generate_id(self, template, name, picture, qrcode, room):
        id = Image.open(template)

        name_pos = (self.NAME_POS[0], self.NAME_POS[1])
        name_size = (self.NAME_POS[2], self.NAME_POS[3])
        id.paste(self.generate_text(name_size, name, ImageFont.truetype(self.FONT, self.NAME_SIZE), self.NAME_COLOR), name_pos)
        
        if room != None:
            room_pos = (self.RN_POS[0], self.RN_POS[1])
            room_size = (self.RN_POS[2], self.RN_POS[3])
            id.paste(self.generate_text(room_size, room, ImageFont.truetype(self.FONT, self.RN_SIZE), self.RN_COLOR), room_pos)

        if picture != None:
            pic_pos = (self.PICTURE_POS[0], self.PICTURE_POS[1])
            pic_size = (self.PICTURE_POS[2], self.PICTURE_POS[3])
            if self.ASPECT_RATIO.lower() == 'stretch':
                id.paste(picture.resize(pic_size), pic_pos)
            else: #keep
                resized_pic = ImageOps.contain(picture, pic_size)
                new_pos = (int((pic_size[0]/2)-(resized_pic.size[0]/2)+pic_pos[0]),
                        int((pic_size[1]/2)-(resized_pic.size[1]/2)+pic_pos[1]))
                id.paste(resized_pic, new_pos)
        elif self.MISSING_PICTURE.lower() == 'skip':
            return

        if qrcode != None:
            qr_pos = (self.QR_POS[0], self.QR_POS[1])
            qr_size = (self.QR_POS[2], self.QR_POS[3])
            id.paste(qrcode.resize(qr_size), qr_pos)

        return id


    #generate a new clean excel file ready for use
    def new_excel(self):
        Wb = Workbook()
        sheet = Wb.worksheets[0]
        sheet.column_dimensions[self.NAME].width = 25
        sheet.column_dimensions[self.PICTURE].width = 50
        
        for i in range(200):
            sheet.row_dimensions[i+1].height = 200

        Wb.save(self.EXCEL)


    #clear excel file
    def clear_excel(self):
        os.remove(self.EXCEL)
        self.new_excel()


    #generate ids from excel file
    def generate_ids(self):
        if not os.path.isfile(self.EXCEL):
            self.new_excel()
            print("Excel File not found. Generating a new excel file.")
            return

        print("Extracting Names")
        names = self.compile_names(self.EXCEL)

        rooms = None
        if self.ROOM_NUMBER != '0':
            rooms = self.compile_rooms(self.EXCEL)

        pictures = self.load_images(self.EXCEL, self.PICTURE, len(names))

        if self.END_ROW > 0:
            names = names[:self.END_ROW]
        if self.START_ROW > 0:
            names = names[self.START_ROW-1:]
            pictures = pictures [self.START_ROW-1:]
            if self.ROOM_NUMBER != '0':
                rooms = rooms[self.START_ROW-1:]

        str_names = "\n".join(names)

        print("Generating QR Codes")
        qrcodes = SimpleQR("\n".join(names), self.LINK).generate(invert=self.INVERT, split=self.SPLIT, save=False, clear=self.DELETE_PREV, border=self.BORDER_SIZE, room=rooms)

        print("Generating IDs")
        skipped = []
        if not os.path.isdir("temp"):
            os.makedirs("temp")
        for i in range(len(names)):
            print("Generating IDs ({}/{})".format(i+1, len(names)))

            room = ''
            if rooms != None:
                room = rooms[i]

            id = self.generate_id(self.TEMPLATE, names[i].upper(), pictures[i], qrcodes[i], None)#, room)
            if id == None:
                skipped.append(names[i])
            else:
                id.save('temp/'+names[i]+'.png')
        
        print("Missing Pictures:\n" + '\n'.join(skipped) + '\n')
        print("Done! (" + str(len(names)-len(skipped)) + "/" + str(len(names)) + ") have been generated successfully.")
        webbrowser.open('file://' + os.getcwd().replace('\\', '/') + '/temp')
        
        if self.CONFIRM:
            inp = input("Confirm QR Codes? y/enter=yes,r=retry,n=cancel \n")
            if inp.lower() == 'r':
                print("Retrying...")
                self.generate_ids()
                return
            elif inp.lower() == 'n':
                print("Cacnceled!")
                shutil.rmtree("temp")
                return
            elif inp.lower() in ['', 'y']:
                print("Confirmed!")

        if self.DELETE_PREV:
            shutil.rmtree("exports")
            os.makedirs("exports")
        for file in os.listdir("temp"):
            shutil.copyfile("temp/"+file,"exports/"+file)
        shutil.rmtree("temp")

        if self.CLEAR_EXCEL:
            self.clear_excel()
        webbrowser.open('file://' + os.getcwd().replace('\\', '/') + '/exports')