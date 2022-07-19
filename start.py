import argparse
from argparse import RawTextHelpFormatter
import sys as _sys
from main import Config
import subprocess


#custom help message
class MyArgumentParser(argparse.ArgumentParser):
    def print_help(self, file=None):
        if file is None:
            file = _sys.stdout
        message = """
enter any of the following commands:

                                leave empty to open gui
'bitly'                         display if bitly in enabled
'bitly -c t'                    enable bitly
'bitly -c f'                    disable bitly
'border_size'                   display box size
'border_size -c <value>'        change box size
'box_size'                      display border size
'box_size -c <value>'           change border size
'cell'                          display starting cell
'cell -c <cell>'                change starting cell
'clear'                         clear config file
'code'                          display all g forms codes
'code -a <code>=<column>'       add g forms code
'code -a <cde>=<clm>+<clm>+etc' add multiple g forms code
'code -r <code>'                remove code
'output'                        display output file path
'output -c <path>'              change output file path
'equal'                         display if filter_equal is enabled
'equal -c t'                    enable filter_equal
'equal -c f'                    disable filter_equal
'excel'                         display excel file path
'excel -c <file>'               change excel file path
'form'                          display form link
'form -c "<url>" '              change form link
'gen' or 'generate'             generate qr codes
'invert'                        display if invert colors in enabled
'invert -c t'                   enable invert colors
'invert -c f'                   disable invert colors
'token'                         display bitly tokens and token usage
'token -a <token>'              add bitly token
'token -a <token>,<token>,etc'  add bitly tokens
'token -r <token>'              remove bitly token
'token -r <token>,<token>,etc'  remove bitly tokens
'token -r all'                  remove all bitly token

    """
        file.write(message+"\n")

#create instance of config
config_instance = Config()

#argparse
parser = MyArgumentParser(
    prog="prefillqr",
    formatter_class=RawTextHelpFormatter
)
parser.add_argument(
    "command",
    choices=['gen','generate','excel','form','token','output','run','invert','box_size','border_size', 'cell', 'code', 'clear', 'bitly', 'equal'],
    nargs='?',
    default='run'
)
parser.add_argument("-c","--change")
parser.add_argument("-a","--add")
parser.add_argument("-r","--remove")
args = parser.parse_args()



#clause guards for unexpected arguments

#using -c in token and code
if args.change is not None and (args.command in ['token', 'code']):
        print("Error: unexpected use of argument. use -a or -r")
        exit()

#using -a and -r unexpectedly
if (args.add is not None or args.remove is not None) and args.command in ['excel','output','form', 'cell', 'invert', 'box_size', 'border_size', 'bitly']:
        print("Error: unexpected use of argument. use -c instead")
        exit()

#using arguments in commands that does not use them
if not(args.add is None and args.remove is None and args.change is None): 
    if args.command == 'run':
        print("Error: You must type a command first before using arguments")
        exit()
    elif args.command == 'gen' or args.command == 'generate':
        print("Error: gen/generate command does not accept arguments")
        exit()
    elif args.command == 'clear':
        print("Error: clear command does not accept arguments")
        exit()

#making sure -c in box size and border size is an integer from 0-100
if args.command in ['box_size', 'border_size']:
    if args.change != None:
        try:
            int(args.change)
        except:
            print("Error: box_size and border_size only accepts integers as values")
            exit()
        if not(int(args.change) > 0):
            print("Error: box_size and border_size must have a value greater than 0")
            exit()
        if not(int(args.change) <= 100):
            print("Error: box_size and border_size must have a value less than or equal to 100")
            exit()



#commands

#run command
if args.command == 'run':
    try: #check if tkinter is installed
        import tkinter
        try:
            subprocess.run(['python3', 'gui.py'])
        except:
            subprocess.run(['python', 'gui.py'])
    except:
        print("Error: Tkinter is not installed, GUI is unavailable. You can only use the CLI. try typing \"-h\" for help")

#generate command
elif args.command == 'gen' or args.command == 'generate':
        config_instance.run()

#excel command
elif args.command == 'excel':
    if args.change is not None: #change excel file
        config_instance.save('excel_file', args.change, True)
        
    else: #read excel file
        print(config_instance.read('excel_file'))

#output command
elif args.command == 'output': #change output path
    if args.change is not None:
        config_instance.save('output_path', args.change, True)
        
    else: #read output
        print(config_instance.read('output_path'))

#form command
elif args.command == 'form': #change form link
    if args.change is not None:
        config_instance.save('forms_link', args.change, True)
        
    else: #read form link
        print(config_instance.read('forms_link'))

#token command
elif args.command == 'token':
    if args.add is not None: #add token
        config_instance.token_change('add', args.add, True)
        
    if args.remove is not None: 
        if args.remove == 'all': #rclear tokens
            config_instance.token_change('clear', 0, True)
            
        else: #remove token
            config_instance.token_change('remove', args.remove, True)
            
    elif args.add is None: #read token and usage
        print(str(config_instance.read('bitly_token')).translate({ord(c): None for c in '][,'})+'\n'+ config_instance.get_usage())

#invert command
elif args.command == 'invert':
    if args.change is not None: #change invert
        if args.change == 't' or args.change == 'true':
            config_instance.save('invert_color', True, True)
            
        elif args.change == 'f' or args.change == 'false':
            config_instance.save('invert_color', False, True)
            
        else: #if invert is set to something that is not t ot f
            print("Error: invert -c only accepts t/true or f/false. Type \"-h\" for more help")
            
    else: #read invert
        print(config_instance.read('invert_color'))

#equal command
elif args.command == 'equal':
    if args.change is not None: #change equal
        if args.change == 't' or args.change == 'true':
            config_instance.save('filter_equal', True, True)
            
        elif args.change == 'f' or args.change == 'false':
            config_instance.save('filter_equal', False, True)
            
        else: #if invert is set to something that is not t ot f
            print("Error: equal -c only accepts t/true or f/false. Type \"-h\" for more help")
            
    else: #read invert
        print(config_instance.read('filter_equal'))

#bitly command
elif args.command == 'bitly': 
    if args.change is not None: #change bitly
        if args.change == 't' or args.change == 'true':
            config_instance.save('use_bitly', True, True)
            
        elif args.change == 'f' or args.change == 'false':
            config_instance.save('use_bitly', False, True)
            
        else: #if bitly is set to something that is not t ot f
            print("Error: bitly -c only accepts t/true or f/false. Type \"-h\" for more help")
            
    else: #read bitly
        print(config_instance.read('use_bitly'))

#box size command
elif args.command == 'box_size':
    if args.change is not None: #change box size
        config_instance.save('box_size', args.change, True)
        
    else: #read box size
        print(config_instance.read('box_size'))

#border size command
elif args.command == 'border_size':
    if args.change is not None: #change border size
        config_instance.save('boreder_size', args.change, True)
        
    else: #read border size
        print(config_instance.read('border_size'))

#cell command
elif args.command == 'cell':
    if args.change is not None: #change starting cell
        config_instance.save('starting_cell', args.change, True)
        
    else: #read starting cell
        print(config_instance.read('starting_cell'))

#code comand
elif args.command == 'code':
    if args.add is not None: #add code
        a = args.add.split('=')
        config_instance.code_change('add', key=a[0], val=a[1])
        
    if args.remove is not None: #remove code
        config_instance.code_change('remove', key=args.remove)
        
    elif args.add is None: #read code
        print(config_instance.read('code'))

#clear command
elif args.command == 'clear':
    config_instance.delete_config_instance #delete cfg.ini
        
else:
    print('Unknown command, type \"-h\" for help')
