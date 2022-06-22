from configparser import ConfigParser
import pandas as pd
import bitlyshortener
import qrcode
from PIL import Image
import os
import re

class Config:
    cfg = ConfigParser()
    excel_file = ""
    forms_link = ""
    destination = "exports/"
    bitly_token = []
    invert_color = False
    progress_bar = None
    output_text = None
    starting_cell = 'A1'
    box_size = 10
    border_size = 4
    code = {}
    use_bitly = False
    
    #pass progress bar and output text object
    def assign_progress_bar(self, bar, text):
        self.progress_bar = bar
        self.output_text = text
    
    def delete_config(self):
        os.remove('cfg.ini')
    
    #read config data
    def read(self, key):
        val = ''
        self.cfg.read('cfg.ini')
        if not self.cfg.has_section('main'):
            self.cfg.add_section('main')

        if self.cfg.has_option('main','excel_file') and (key == 'all' or key == 'excel_file'):
            val = self.cfg.get('main', 'excel_file')
            self.excel_file = val

        if self.cfg.has_option('main','forms_link') and (key == 'all' or key == 'forms_link'):
            val = self.cfg.get('main', 'forms_link')
            self.forms_link = val

        if self.cfg.has_option('main','destination') and (key == 'all' or key == 'destination'):
            val = self.cfg.get('main', 'destination')
            self.destination = val

        if self.cfg.has_option('main','bitly_token') and (key == 'all' or key == 'bitly_token'):
            val = self.cfg.get('main', 'bitly_token').split(",")
            val = list(filter(None, val))
            self.bitly_token = val
        if self.cfg.has_option('main', 'invert_color') and (key == 'all' or key == 'invert_color'):
            strval = self.cfg.get('main','invert_color')
            if strval == 'True':
                val = True
            else:
                val = False
            self.invert_color = val
        if self.cfg.has_option('main', 'use_bitly') and (key == 'all' or key == 'use_bitly'):
            strval = self.cfg.get('main','use_bitly')
            if strval == 'True':
                val = True
            else:
                val = False
            self.use_bitly = val
        if self.cfg.has_option('main','starting_cell') and (key == 'all' or key == 'starting_cell'):
            val = self.cfg.get('main', 'starting_cell')
            self.starting_cell = val
        if self.cfg.has_option('main','box_size') and (key == 'all' or key == 'box_size'):
            val = self.cfg.get('main', 'box_size')
            self.box_size = int(val)
        if self.cfg.has_option('main','border_size') and (key == 'all' or key == 'border_size'):
            val = self.cfg.get('main', 'border_size')
            self.border_size = int(val)
        if self.cfg.has_option('main','code') and (key == 'all' or key == 'code'):
            val = self.cfg.get('main', 'code')
            if ' = ' in val:
                val = dict(item.split(" = ") for item in val.split(", "))
                self.code = val

        if key != 'all':
            return val
        else:
            return {
                'excel_file': self.excel_file,
                'forms_link': self.forms_link,
                'destination': self.destination,
                'bitly_token': self.bitly_token,
                'invert_color': self.invert_color,
                'starting_cell': self.starting_cell,
                'box_size': self.box_size,
                'border_size': self.border_size,
                'code': self.code,
                'use_bitly': self.use_bitly
            }
    

    #save data
    def save(self, key, val, prnt):
        if key == "bitly_token":
            val = ",".join(val)
            
        else:
            self.read('all')
            val = val
        if not self.cfg.has_section('main'):
            self.cfg.add_section('main')
        self.cfg.read('cfg.ini')
        self.cfg.set('main', str(key), str(val))
        with open('cfg.ini', 'w') as f:
            self.cfg.write(f)
        if prnt:
            print('save ' + str(key) + ' as ' + str(val))

            
    def code_change(self, cmd, **kwargs):
        key = kwargs.get('key', None)
        val = kwargs.get('val', None)
        self.read('all')
        
        if cmd == 'add':
            self.code[key] = val 
        if cmd == 'remove':
            del self.code[key]
        
        code_str = str(self.code).replace(' ','').replace('{','').replace('}','').replace(':', ' = ').replace(',', ', ').replace('\'', ''). replace('\"', '')
        print(code_str)
        
        self.save('code', code_str, True)
        
    
    
            
    #change token to val
    def token_change(self, cmd, val, prnt):
        self.read('all')
        new_tokens = self.bitly_token
        if cmd == "clear":
            self.save('bitly_token', [''], prnt)
            self.get_usage()

        if cmd == "add":
            if "," in val:
                val
                tokens = val.replace(" ", "").split(',')
                new_tokens = [i for n, i in enumerate(tokens) if i not in tokens[:n]]
                new_tokens = list(filter(None, new_tokens))
                
                self.save('bitly_token',new_tokens, prnt)
            elif not val in new_tokens:
                tokens = val.replace(" ", "")
                new_tokens.append(val)
                if new_tokens == '':
                    return
                self.save('bitly_token', new_tokens, prnt)
            self.get_usage()

        if cmd == "remove":
            if "," in val:
                tokens = val.split(",")
                tokens = list(set(tokens).intersection(new_tokens))
                for element in tokens:
                    if element in new_tokens:
                        new_tokens.remove(element)
                new_tokens = list(filter(None, new_tokens))
                self.save('bitly_token',new_tokens, prnt)
            elif val in new_tokens:
                new_tokens.remove(val)
                self.save('bitly_token',new_tokens, prnt)
            self.get_usage()

    def get_usage(self):
        if self.read('use_bitly') and self.read('bitly_token') != []:
            try:
                shortener = bitlyshortener.Shortener(tokens=self.read('bitly_token'))
                return str('Tokens Usage: '+str(shortener.usage()*100)[0:5]+'%')
            except:
                return 'Tokens Usage: Error'

                
    def run(self, **kwargs):
        if self.progress_bar is None:
            Generate(self.read('all'), None, None)
        else:
            Generate(self.read('all'), self.progress_bar, self.output_text)
            usg = kwargs.get('usg', None)
            if usg != None:
                usage = self.get_usage()
                usg.config(text=usage)

        
        
        
        
        
        
        
        
class Generate():
    def __init__(self, cfg, progress_bar, output_text):
        
        self.progress_bar = progress_bar
        self.output_text = output_text
        self.progress('start', 'Starting...')

        #check for missing values
        if cfg['excel_file'] == '':
            print("Error: No excel file found! Make sure you have entered the excel file correctly.")
            if not progress_bar is None:
                output_text.config(text="Error: No excel file found!")
            return
        if cfg['forms_link'] == '':
            print("Error: G forms link not found! Make sure you have entered the link correctly.")
            if not progress_bar is None:
                output_text.config(text="Error: No forms link entered")
            return
        if cfg['bitly_token'] == [] and cfg['use_bitly']:
            print("Error: No bitly tokens found! Make sure you have entered the tokens correctly.")
            if not progress_bar is None:
                output_text.config(text="Error: No bitly tokens added")
            return
        if not 'filename' in cfg['code']:
            print("Error: no code \'filename\' found! add code \'filename\' first in order to determine the filenames of the output files")
            if not progress_bar is None:
                output_text.config(text="Error: code \'Filename\' not found")
            return
        
        #set data frame with starting cell
        c_and_r = re.split('(\d+)',cfg['starting_cell'])
        c = self.col2num(c_and_r[0].upper())
        r = int(c_and_r[1]) - 1
        
        df = pd.read_excel(pd.ExcelFile(cfg['excel_file']), skiprows=r, header=None)
        code = cfg['code']
        if c > 0:
            for i in range(c):
                del df[df.columns[0]]
                
        #parse filenames
        self.progress(0, 'Parsing files...')
        filenames = self.parse_filenames(df, code['filename'])
        self.ammount = len(filenames)
        if cfg['use_bitly']:
            self.total = self.ammount*4
        else:
            self.total = self.ammount*3
        self.progress(self.ammount, 'Generating prefill links...')
        
        #generate prefill links
        urls = self.generate_links(cfg['forms_link'], df, code, filenames, c)
        
        #shorten links
        if cfg['use_bitly']:
            self.progress(0, 'Shortening Links...')
            urls = self.shorten(cfg['bitly_token'], urls)
            self.progress(self.ammount, 'Generating qr codes...')
        
        #generate qr codes
        for n in range(len(filenames)):
            self.progress(1, 'Generating qr code ('+str(n+1)+'/'+str(self.ammount)+')')
            self.generate_qr(cfg['destination'], urls[n], filenames[n], cfg['invert_color'], cfg['box_size'], cfg['border_size'])
            
        self.progress('done', 'Done!')

    def col2num(self, col):
        c = 0;
        for b in range(len(col)):
            c *= 26;
            c += ord(col[b]) - ord('A') + 1;
        c -= 1
        return c

    #parse data frame and get filenames
    def parse_filenames(self, df, code):
        filenames = []
        fn_columns = code.upper().split('+')
        for letter in range(len(fn_columns)):
            fn_columns[letter] = self.col2num(fn_columns[letter].upper())
        fn_columns = list(filter((-22).__ne__, fn_columns))

        fn_list = list(df.iloc[:, fn_columns[0]])
        
        #find how many qr codes to be generated
        length = 0
        for row in range(len(fn_list)+1):
            length = row
            stop = False
            try:
                if pd.isnull(fn_list[row]):
                    stop = True
            except:
                break
            if stop == True:
                break
        
        #find and remove rows after the length
        length = len(fn_list) - length
        for i in range(length):
            del fn_list[-1]
            
        #combine columns if required
        if len(fn_columns) > 1:
            fn = fn_list
            filenames = fn
            for rc in range(len(fn_columns)-1):
                fn = ['{} '.format(elem) for elem in filenames]
                filenames = [i + j for i, j in zip(fn, list(df.iloc[:, fn_columns[rc+1]]))]   
        else:
            filenames = fn_list
        return filenames

    
    #generate links from data frame
    def generate_links(self, forms_link, df, code, fn, c):
        urls = []
        for n in range(len(fn)):
            self.progress(1, 'Generating prefill links ('+str(n+1)+'/'+str(self.ammount)+')')
            link = forms_link
            for cd in code.keys():
                if cd != 'filename':
                    cd_list = code[cd].upper().split('+')
                    ans = ''
                    for cl in cd_list:
                        
                        ncl = self.col2num(cl)
                        ans += str(df.iloc[n, ncl]) + ' '
                        
                    link = link.replace('='+cd,'='+ans.strip())
            urls.append(link)
        return urls

    
    #shorten links with bitly
    def shorten(self, bitly_token,urls): 
        shortener = bitlyshortener.Shortener(tokens=bitly_token)
        urls= shortener.shorten_urls(urls)
        return urls


    #generates an image of a qr code that links to specified url
    def generate_qr(self, destination, url, name, invert, box, border):
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
        path = destination
        if path == '':
            path = 'exports'
        if not os.path.exists(path): #generate path
            if path[-1] == "/":
                path = path[:-1]
            os.makedirs(path)
        if not path[-1] == "/":
            path = path + "/"
        path = path + str(name) + ".png"

        img.save(path) #save

    def progress(self, progress, text):
        print(text)
        if not self.progress_bar is None:
            self.output_text.config(text=text)
            if str(progress) == 'start':
                self.progress_bar.grid(column=0, row=1, padx=20, pady=5, sticky='ew')
                self.progress_bar['value'] = 0
            elif str(progress) == 'done':
                self.progress_bar['value'] = 100
            elif progress > 0:
                self.progress_bar['value'] += (progress/(self.total))*100
            
            