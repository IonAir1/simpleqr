from configparser import ConfigParser
from PIL import Image
import os
import qrcode


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
        for f in os.listdir(self.output_path):
            if not f.endswith(".png"):
                continue
            os.remove(os.path.join(self.output_path, f))
        
        #generate qr codes
        a = (1/len(filenames))*100
        qrcodes = []
        for n in range(len(filenames)):
            qrcodes.append(self.generate_qr(self.output_path, urls[n], filenames[n], kwargs.get('invert', False), kwargs.get('box', 10), kwargs.get('border', 4)))
            self.progress += a

        self.progress = 100
        return qrcodes
    
    
    #generate links
    def generate_links(self, names, url, **kwargs):
        if kwargs.get('replace', True):
            urls = []
            for name in names:
                if kwargs.get('split', False) == '1': #check if split name is on
                    last_first = name.split(',')
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
                            final_url = final_url.replace('addfirstname','').replace('addmiddlename','').replace('firstname','').replace('middlename','')
                            urls.append(final_url)

                        else:
                            new_url = new_url.replace('=firstname','='+last_first[1].strip().replace(' ', '%20'))
                            urls.append(new_url.replace('firstname','').replace('middlename',''))
                    else:
                        urls.append(new_url.replace('firstname','').replace('middlename',''))

                else:
                    final_url = url.replace('=name','='+name.replace(' ', '%20'))
                    urls.append(final_url)

            return urls
        elif not(kwargs.get('replace', False)):
            return names


    #generates an image of a qr code that links to specified url
    def generate_qr(self, output_path, url, name, invert, box, border):
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
        img.save(output_path.replace('\\','/') + str(name).replace('/','').replace(':','').replace('.', '') + ".png") #save
        return img
            