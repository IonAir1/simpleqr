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
'cell'                          display starting cell
'cell -c <cell>'                change starting cell
'code'                          display all g forms codes
'code -a <code>=<column>'       add g forms code
'code -a <cde>=<clm>+<clm>+etc' add multiple g forms code
'code -r <code>'                remove code
'destination'                   display destination file path
'destination -c <path>'         change destination file path
'excel'                         display excel file path
'excel -c <file>'               change excel file path
'form'                          display form link
'form -c "<url>" '              change form link
'gen' or 'generate'             generate qr codes
'invert'                        show if invert colors in enabled
'invert -c t'                   enable invert colors
'invert -c f'                   disable invert colors
'token'                         display bitly tokens
'token -a <token>'              add bitly token
'token -a <token>,<token>,etc'  add bitly tokens
'token -r <token>'              remove bitly token
'token -r <token>,<token>,etc'  remove bitly tokens
'token -r all'                  remove all bitly token
'usage                          display usage of bitly tokens

    """
        file.write(message+"\n")
        
#argparse
parser = MyArgumentParser(
    prog="prefillqr",
    formatter_class=RawTextHelpFormatter
)
parser.add_argument(
    "command",
    choices=['gen','generate','excel','form','token','destination','run','invert','box_size','border_size', 'cell', 'code', 'usage'],
    nargs='?',
    default='run'
)
parser.add_argument("-c","--change")
parser.add_argument("-a","--add")
parser.add_argument("-r","--remove")
args = parser.parse_args()

if args.command == 'run':
    try:
        subprocess.run(['python3', 'gui.py'])
    except:
        subprocess.run(['python', 'gui.py'])
elif args.command == 'gen' or args.command == 'generate':
    if args.add is None and args.remove is None and args.change is None:
        Config().run()
    else:
        print("gen/generate command does not accept arguments")
elif args.command == 'excel':
    if args.change is not None:
        Config().save('excel_file', args.change, True)
    else:
        print(Config().read('excel_file'))
elif args.command == 'destination':
    if args.change is not None:
        Config().save('destination', args.change, True)
    else:
        print(Config().read('destination'))
elif args.command == 'form':
    if args.change is not None:
        Config().save('forms_link', args.change, True)
    else:
        print(Config().read('forms_link'))
elif args.command == 'token':
    if args.add is not None:
        Config().token_change('add', args.add, True)
    if args.remove is not None:
        if args.remove == 'all':
            Config().token_change('clear', 0, True)
        else:
            Config().token_change('remove', args.remove, True)
    elif args.add is None:
        print(str(Config().read('bitly_token')).translate({ord(c): None for c in '][,'}))
elif args.command == 'invert':
    if args.change is not None:
        if args.change == 't' or args.change == 'true':
            Config().save('invert_color', True, True)
        elif args.change == 'f' or args.change == 'false':
            Config().save('invert_color', False, True)
    else:
        print(Config().read('invert_color'))

elif args.command == 'box_size':
    if args.change is not None:
        Config().save('box_size', args.change, True)
    else:
        print(Config().read('box_size'))

elif args.command == 'border_size':
    if args.change is not None:
        Config().save('boreder_size', args.change, True)
    else:
        print(Config().read('border_size'))
elif args.command == 'cell':
    if args.change is not None:
        Config().save('starting_cell', args.change, True)
    else:
        print(Config().read('starting_cell'))
elif args.command == 'code':
    if args.add is not None:
        
        a = args.add.split('=')
        Config().code_change('add', key=a[0], val=a[1])
    if args.remove is not None:
        Config().code_change('remove', key=args.remove)
    elif args.add is None:
        print(Config().read('code'))
elif args.command == 'usage':
    print(Config().get_usage())
        
else:
    print('Unknown command, type \"-h\" for help')
if args.change is not None and args.command == "token":
        print("unexpected use of argument. Use -a or -r for tokens")
if (args.add is not None or args.remove is not None) and args.command in ['excel','destination','form']:
        print("unexpected use of argument. Use -c instead")