from main import Config
import os
import threading
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd


#root window
root = tk.Tk()
root.title('PrefillQR')
root.geometry('900x500+50+50')
root.resizable(False, False)
config_instance = Config()


#initializing variables
data = config_instance.read('all')
bitly_token = data['bitly_token']
ef_var = tk.StringVar(root, data['excel_file'])
fl_var = tk.StringVar(root, data['forms_link'])
df_var = tk.StringVar(root, data['destination'])
bs_var = tk.IntVar(root, data['box_size'])
br_var = tk.IntVar(root, data['border_size'])
ic_var = tk.BooleanVar(root, data['invert_color'])


#open window for selecting excel file
def select_excel_file():
    filetypes = (
        ('Excel files', '*.xlsx'),
        ('All files', '*.*')
    )
    filename = fd.askopenfilename(
        title='Open a file',
        initialdir=os.path.expanduser('~'),
        filetypes=filetypes)
    ef_var = tk.StringVar(root, filename)
    ef_entry.delete(0,tk.END)
    ef_entry.insert(0,filename)
    config_instance.save('excel_file', filename)

    
#open window for selecting output folder
def select_destination():
    foldername = fd.askdirectory(
        title='Select a folder',
        initialdir=os.path.expanduser('~')
    )
    df_var = tk.StringVar(root, foldername)
    df_entry.delete(0,tk.END)
    df_entry.insert(0,foldername)
    config_instance.save('destination', foldername)

    
#create new thread to start generating qr codes
def generate():
    generate = threading.Thread(target=config_instance.run)
    generate.start()

    
    
    
#title at top
title = ttk.Label(root, text="Create QR Codes to prefilled G Forms")
title.grid(column=0, row=0, padx=10, pady=10, sticky='w')



#excel file section
ef = ttk.LabelFrame(root, text='Excel File') #excel file frame
ef.grid(column=0, row=1, padx=10, pady=10, sticky='w')

ef_entry = ttk.Entry(ef, textvariable=ef_var, width=50) #excel file entry input
ef_entry.grid(column=0, row=0, padx=10, pady=10)
ef_entry.bind("<FocusOut>", lambda event: config_instance.save('excel_file', ef_var.get()))
ef_entry.bind('<Control-a>', lambda x: ef_entry.selection_range(0, 'end') or "break")

ef_browse = ttk.Button(ef, text='Browse', command=select_excel_file) #excel file browse button
ef_browse.grid(column=1, row=0, padx=10, pady=10)


#forms link section
fl = ttk.LabelFrame(root, text='G Forms Prefill link') #forms link frame
fl.grid(column=0, row=2, padx=10, pady=10, sticky='w')

fl_entry = ttk.Entry(fl, textvariable=fl_var, width=50) #forms link entry input
fl_entry.bind("<FocusOut>", lambda event: config_instance.save('forms_link', fl_var.get()))
fl_entry.bind('<Control-a>', lambda x: fl_entry.selection_range(0, 'end') or "break")
fl_entry.grid(column=0, row=0, padx=10, pady=10)



#options section
op = ttk.Frame(root)
op.grid(column=0, row=3, padx=10, pady=10, sticky='w')


bs = ttk.Frame(op)#box size frame
bs.grid(column=0, row=0,padx=30, pady=10, sticky='w')
bs.grid_columnconfigure(0, weight=1)

bs_spinbox = ttk.Spinbox(bs, textvariable=bs_var, from_=0, to=100, width=3) #box size spinbox
bs_spinbox.bind("<FocusOut>", lambda event: config_instance.save('box_size', bs_var.get()))
bs_spinbox.grid(column=0, row=0)

bs_text = ttk.Label(bs, text='Box Size') #box size label
bs_text.grid(column=1, row=0)


br = ttk.Frame(op) #border size frame
br.grid(column=1, row=0,padx=20, pady=10, sticky='w')
br.grid_columnconfigure(0, weight=1)

br_spinbox = ttk.Spinbox(br, textvariable=br_var, from_=0, to=100, width=3) #border size spinbox
br_spinbox.bind("<FocusOut>", lambda event: config_instance.save('border_size', br_var.get()))
br_spinbox.grid(column=0, row=0)

br_text = ttk.Label(br, text='Border Size') #border size label
br_text.grid(column=1, row=0)


#invert color checkbox
ic = ttk.Checkbutton(op,
                text='Invert Color',
                command=lambda: config_instance.save('invert_color', ic_var.get()),
                variable=ic_var,
                onvalue=True,
                offvalue=False)
ic.grid(column=2, row=0, padx=20, pady=10)



#destination folder section
df = ttk.LabelFrame(root, text='Destination Folder') #df frame
df.grid(column=0, row=4, padx=10, pady=10, sticky='w')

df_entry = ttk.Entry(df, textvariable=df_var, width=50) #df entry input
df_entry.grid(column=0, row=0, padx=10, pady=10)
df_entry.bind("<FocusOut>", lambda event: config_instance.save('destination', df_var.get()))
df_entry.bind('<Control-a>', lambda x: df_entry.selection_range(0, 'end') or "break")

df_browse = ttk.Button(df, text='Browse', command=select_destination) #df browse button
df_browse.grid(column=1, row=0, padx=10, pady=10)



#seperator
separator = ttk.Separator(root, orient='vertical')
separator.grid(column=1, row=0, rowspan=5, sticky='ns')



#generate button
gn = ttk.Button(root,
                text='Generate',
                command=generate)
gn.grid(column=2, row=5, sticky='se', padx=30)



#progress section
ps = ttk.Frame(root) #progress section text
ps.grid(column=0, row=6, columnspan=3, sticky='ew')
ps.columnconfigure(0, weight=1)

pt = ttk.Label(ps, text='') #progress text
pt.grid(column=0, row=0, padx=30, sticky='w')

pb = ttk.Progressbar( #progress bar
    ps,
    orient='horizontal',
    mode='determinate',
    length=480,
)
config_instance.assign_progress_bar(pb, pt)



root.mainloop()