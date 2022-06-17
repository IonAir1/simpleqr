from configparser import ConfigParser
import prefill

class Config:
    cfg = ConfigParser()
    excel_file = ""
    forms_link = ""
    destination = "exports/"
    bitly_token = []
    invert_color = False
    progress_bar = None
    output_text = None
    
    #pass progress bar and output text object
    def assign_progress_bar(self, bar, text):
        self.progress_bar = bar
        self.output_text = text
    
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
            
            
        if key != 'all':
            return val
        else:
            return {
                'excel_file': self.excel_file,
                'forms_link': self.forms_link,
                'destination': self.destination,
                'bitly_token': self.bitly_token,
                'invert_color': self.invert_color
            }
    

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
        self.cfg.set('main', str(key), str(val))
        with open('cfg.ini', 'w') as f:
            self.cfg.write(f)
        print('save ' + str(key) + ' as ' + str(val))

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
            'bitly_token': self.bitly_token,
            'invert_color': self.invert_color
            }
        Generate(cfg, self.progress_bar, self.output_text)

        
        
        
        
        
        
        
        
class Generate():
    def __init__(self, cfg, progress_bar, output_text):
        #initiate progress bar and output text
        if not progress_bar is None:
            output_text.config(text="Starting...")
        
        #check for missing values
        if cfg['excel_file'] == '':
            print("No excel file found! Make sure you have entered the excel file correctly.")
            if not progress_bar is None:
                output_text.config(text="No excel file found!")
            return
        if cfg['forms_link'] == '':
            print("G forms link not found! Make sure you have entered the link correctly.")
            if not progress_bar is None:
                output_text.config(text="No forms link entered")
            return
        if cfg['bitly_token'] == []:
            print("No bitly tokens found! Make sure you have entered the tokens correctly.")
            if not progress_bar is None:
                output_text.config(text="No bitly tokens added")
            return
        
        
        #read data
        data = prefill.read_file(cfg['excel_file'])
        urls = []
        ammount = len(data)
        total = len(data)*4
        if not progress_bar is None:
            progress_bar.grid(column=0, row=1, padx=20, pady=5, sticky='ew')
            progress_bar['value'] = (ammount/total)*100
    
    
        for i in range(len(data.keys())): #generate prefilled link
            print("generating prefilled link ("+str(i+1)+"/"+str(len(data.keys()))+")")
            if not progress_bar is None:
                progress_bar['value'] += (1/total)*100
                output_text.config(text="generating prefilled link ("+str(i+1)+"/"+str(len(data.keys()))+")") 
                
            item = list(data)[i]
            last_name = data[item][0]
            first_name = data[item][1]
            email = data[item][2]
            url = prefill.generate_prefill(cfg['forms_link'],str(first_name) + " " + str(last_name), str(email))
            data[item].append(url)
            urls.append(url)

        print("Shortening urls...")
        if not progress_bar is None:
            progress_bar['value'] += ammount
            output_text.config(text="Shortening urls...")

        
        urls = prefill.shorten(cfg['bitly_token'],urls) #shorten links

        for i in range(len(data.keys())): #generate qr code
            progress_bar['value'] += (1/total)*100
            print("generating qr code ("+str(i+1)+"/"+str(len(data.keys()))+")")
            if not progress_bar is None:
                output_text.config(text="generating qr code ("+str(i+1)+"/"+str(len(data.keys()))+")")
                               
            item = list(data)[i]
            data[item].append(urls[i])
            prefill.generate_qr(cfg['destination'],urls[i], str(data[item][0]), cfg['invert_color'])

        print("Done!")
        if not output_text is None:
            progress_bar['value'] = 100
            output_text.config(text="Done!")