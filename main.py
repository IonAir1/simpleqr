from configparser import ConfigParser
import prefill

class Config:
    cfg = ConfigParser()
    excel_file = ""
    forms_link = ""
    destination = "exports/"
    bitly_token = []

    
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
        if key != 'all':
            return val
    

    #save data
    def save(self, key, val):
        if key == "bitly_token":
            val = ",".join(val)
        else:
            self.read('all')
            val = val
        if not self.cfg.has_section('main'):
            self.cfg.add_section('main')
        self.cfg.read('cfg.ini')
        self.cfg.set('main', str(key), val)
        with open('cfg.ini', 'w') as f:
            self.cfg.write(f)

    #change token to val
    def token_change(self, cmd, val):
        self.read('all')
        new_tokens = self.bitly_token
        if cmd == "clear":
            self.save('bitly_token', [''])

        if cmd == "add":
            if "," in val:
                tokens = val.split(",")
                new_tokens += [x for x in tokens if x not in new_tokens]
                new_tokens = list(filter(None, new_tokens))
                self.save('bitly_token',new_tokens)
            elif not val in new_tokens:
                new_tokens.append(val)
                self.save('bitly_token', new_tokens)

        if cmd == "remove":
            if "," in val:
                tokens = val.split(",")
                tokens = list(set(tokens).intersection(new_tokens))
                for element in tokens:
                    if element in new_tokens:
                        new_tokens.remove(element)
                new_tokens = list(filter(None, new_tokens))
                self.save('bitly_token',new_tokens)
            elif val in new_tokens:
                new_tokens.remove(val)
                self.save('bitly_token',new_tokens)

    def run(self):
        self.read('all')
        cfg = {
            'excel_file': self.excel_file,
            'forms_link': self.forms_link,
            'destination': self.destination,
            'bitly_token': self.bitly_token
            }
        Generate(cfg)

class Generate():
    def __init__(self, cfg):
        if cfg['excel_file'] == '':
            print("No excel file found! Make sure you have entered the excel file correctly.")
            return
        if cfg['forms_link'] == '':
            print("G forms link not found! Make sure you have entered the link correctly.")
            return
        if cfg['bitly_token'] == []:
            print("No bitly tokens found! Make sure you have entered the tokens correctly.")
            return
        
        data = prefill.read_file(cfg['excel_file'])
        urls = []

        for i in range(len(data.keys())): #generate prefilled link
            print("generating prefilled link ("+str(i+1)+"/"+str(len(data.keys()))+")")
            item = list(data)[i]
            last_name = data[item][0]
            first_name = data[item][1]
            email = data[item][2]
            url = prefill.generate_prefill(cfg['forms_link'],first_name + " " + last_name, email)
            data[item].append(url)
            urls.append(url)

        print("Shortening urls...")
        
        urls = prefill.shorten(cfg['bitly_token'],urls) #shorten links

        for i in range(len(data.keys())): #generate qr code
            print("generating qr code ("+str(i+1)+"/"+str(len(data.keys()))+")")
            item = list(data)[i]
            data[item].append(urls[i])
            prefill.generate_qr(cfg['destination'],urls[i], data[item][0])

        print("Done!")