from main import Config
import os
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd

data = Config().read('all')
excel_file = data['excel_file']
forms_link = data['forms_link']
destination = data['destination']
bitly_token = data['bitly_token']

root = tk.Tk()
root.title('PrefillQR')
root.geometry('800x600+50+50')
root.resizable(False, False)



def select_excel_file():
    filetypes = (
        ('Excel files', '*.xlsx'),
        ('All files', '*.*')
    )

    filename = fd.askopenfilename(
        title='Open a file',
        initialdir=os.path.expanduser('~'),
        filetypes=filetypes)
    excel_file = filename
    ef_entry.delete(0,tk.END)
    ef_entry.insert(0,filename)
    Config().save('excel_file', filename)

def select_destination():
    foldername = fd.askdirectory(
        title='Select a folder',
        initialdir=os.path.expanduser('~')
    )
    destination = foldername
    df_entry.delete(0,tk.END)
    df_entry.insert(0,foldername)
    Config().save('excel_file', foldername)
    
title = ttk.Label(root, text="Create QR Codes to prefilled G Forms")
title.grid(column=0, row=0, padx=10, pady=10, sticky=tk.W)

alignment_var = tk.StringVar()
alignments = ('Left', 'Center', 'Right')

ef = ttk.LabelFrame(root, text='Excel File')
ef.grid(column=0, row=1, padx=10, pady=10, sticky=tk.W)
ef_entry = ttk.Entry(ef, textvariable=excel_file, width=50)
ef_entry.grid(column=0, row=0, padx=10, pady=10)
ef_entry.bind('<Control-a>', lambda x: ef_entry.selection_range(0, 'end') or "break")
ef_entry.insert(0,excel_file)
ef_browse = ttk.Button(ef, text='Browse', command=select_excel_file)
ef_browse.grid(column=1, row=0, padx=10, pady=10)

fl = ttk.LabelFrame(root, text='G Forms Prefill link')
fl.grid(column=0, row=2, padx=10, pady=10, sticky=tk.W)
fl_entry = ttk.Entry(fl, textvariable=forms_link, width=50, validate="focusout", validatecommand=lambda: Config().save('forms_link', forms_link))
fl_entry.grid(column=0, row=0, padx=10, pady=10)

df = ttk.LabelFrame(root, text='Destination Folder')
df.grid(column=0, row=3, padx=10, pady=10, sticky=tk.W)
df_entry = ttk.Entry(df, textvariable=destination, width=50, validate="focusout", validatecommand=lambda: Config().save('destination', destination))
df_entry.grid(column=0, row=0, padx=10, pady=10)
df_entry.bind('<Control-a>', lambda x: df_entry.selection_range(0, 'end') or "break")
df_entry.insert(0,destination)
df_browse = ttk.Button(df, text='Browse', command=select_destination)
df_browse.grid(column=1, row=0, padx=10, pady=10)


root.mainloop()