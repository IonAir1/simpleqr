from configparser import ConfigParser
from PIL import Image
import os
import qrcode

        
class Generate():
    def __init__(self, text, url, **kwargs):
        
        self.text = text
        self.url = url

        output_path = kwargs.get('path','exports/')


        #parse filenames
        filenames = self.text.split('\n')
        

        #generate prefill links
        urls = self.generate_links(filenames, url)
        

        #clear output folder
        if not os.path.exists(output_path):
            os.makedirs(output_path)
        for f in os.listdir(output_path):
            if not f.endswith(".png"):
                continue
            os.remove(os.path.join(output_path, f))


        #generate qr codes
        for n in range(len(filenames)):
            self.generate_qr(output_path, urls[n], filenames[n], kwargs.get('invert', False), kwargs.get('box', 10), kwargs.get('border', 4))

    
    #generate links from data frame
    def generate_links(self, names, url):
        urls = []
        for name in names:
            final_url = url.replace('=name','='+name.replace(' ', '%20').replace('.', ''))
            urls.append(final_url)
        return urls


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
        img.save(output_path + str(name) + ".png") #save
            