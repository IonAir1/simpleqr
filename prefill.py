import bitlyshortener
import qrcode
from PIL import Image
import pandas as pd
import os



#generates a prefill link with the input
def generate_prefill(link,name, email): 
    url = link.replace('=name', '=' + name)
    url = url.replace('=email', '=' + email)
    return(url)


#shortens urls using bitly
def shorten(bitly_token,urls): 
    shortener = bitlyshortener.Shortener(tokens=bitly_token)
    urls= shortener.shorten_urls(urls)
    return urls


#generates an image of a qr code that links to specified url
def generate_qr(destination,url, name):
    qr = qrcode.QRCode( #qr code properties
    version=1,
    error_correction=qrcode.constants.ERROR_CORRECT_L,
    box_size=10,
    border=4,
    )
    
    qr.add_data(url) #qr code data
    qr.make(fit=True)
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
    path = path + name + ".png"
    
    img.save(path) #save


#reads excel file
def read_file(file):
    items = {}
    n = 0
    empty = False
    df = pd.read_excel(file, header=None)
    
    while not empty: #converts data frame to dictionary
        data = []
        data.append(df.iloc[n][0])
        data.append(df.iloc[n][1])
        data.append(df.iloc[n][2])
        items.update({data[0]: data})
        n += 1
        try: #repeat until blank line
            pd.isnull(df.iloc[n][0])
        except:
            empty = True
    return items

