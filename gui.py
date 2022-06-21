from main import Config
import os
import threading
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd


#root window
root = tk.Tk()
root.title('PrefillQR')
root.geometry('720x480+50+50')
root.resizable(False, False)
config_instance = Config()


#initializing variables
data = config_instance.read('all')
ef_var = tk.StringVar(root, data['excel_file'])
fl_var = tk.StringVar(root, data['forms_link'])
df_var = tk.StringVar(root, data['destination'])
sc_var = tk.StringVar(root, data['starting_cell'])
bs_var = tk.IntVar(root, data['box_size'])
br_var = tk.IntVar(root, data['border_size'])
ic_var = tk.BooleanVar(root, data['invert_color'])
bitly_token = data['bitly_token']
code = data['code']
hidden_token = []
bt_add = tk.StringVar()
bt_show = tk.BooleanVar(root, False)

cd_add_cd = tk.StringVar()
cd_add_rp = tk.StringVar()


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
    config_instance.save('excel_file', filename, True)

    
#open window for selecting output folder
def select_destination():
    foldername = fd.askdirectory(
        title='Select a folder',
        initialdir=os.path.expanduser('~')
    )
    df_var = tk.StringVar(root, foldername)
    df_entry.delete(0,tk.END)
    df_entry.insert(0,foldername)
    config_instance.save('destination', foldername, True)


    
#create new thread to start generating qr codes
def generate():
    gn.focus_set()
    generate = threading.Thread(target=config_instance.run, kwargs={'usg':usg})
    generate.start()



def add_token():
    global bitly_token
    token = bt_add.get().replace(" ", "").split(",")
    tokens = [i for n, i in enumerate(token) if i not in token[:n]]
    bitly_token += [x for x in tokens if x not in bitly_token]
    bitly_token = list(filter(None, bitly_token))
    
    config_instance.token_change('add',bt_add.get(), bt_show.get())
    ba_entry.delete(0,tk.END)
    show_token()

def remove_token():
    selected_iid = bt_tree.selection()
    tokens = []
    for element in selected_iid:
        tokens.append(bitly_token[(bt_tree.index(element))])
    for element in tokens:
        bitly_token.remove(element)
    config_instance.token_change('remove', ','.join(tokens), bt_show.get())
    show_token()

    
def show_token():
    for item in bt_tree.get_children():
      bt_tree.delete(item)
    if bt_show.get():
        ba_entry['show'] = ''
        for element in bitly_token:
            bt_tree.insert('', tk.END, values=element)
    else:
        ba_entry['show'] = '*'
        for element in bitly_token:
            token = '*' * len(element)
            hidden_token.append(token)
            bt_tree.insert('', tk.END, values=token)
    
def hide_token(x):
    bt_show.set(False)
    show_token()
    
def add_code():
    key = cd_add_cd.get()
    val = cd_add_rp.get().upper()
    ca_entry_cd.delete(0,tk.END)
    ca_entry_rp.delete(0,tk.END)
    code[key] = val
    config_instance.code_change('add', key=key, val=val)
    show_code()
    
def remove_code():
    selected_iid = cd_tree.selection()
    codes = []
    for element in selected_iid:
        print(list(code)[cd_tree.index(element)])
        codes.append(list(code)[cd_tree.index(element)])
    for element in codes:
        del code[element]
        config_instance.code_change('remove', key=element)
    show_code()

def show_code():
    for item in cd_tree.get_children():
        cd_tree.delete(item)
    for element in code:
        cd_tree.insert('', tk.END, values=[element, code[element].upper()])

notebook = ttk.Notebook(root)
notebook.pack(fill='x', side='top')
notebook.bind("<<NotebookTabChanged>>", hide_token)
st = ttk.Frame(notebook, width=400, height=280)
cd = ttk.Frame(notebook, width=400, height=280)
bt = ttk.Frame(notebook, width=400, height=280)
st.pack(fill='both', expand=True)
cd.pack(fill='both', expand=True)
bt.pack(fill='both', expand=True)
notebook.add(st, text='Settings')
notebook.add(cd, text='Codes')
notebook.add(bt, text='Bitly Tokens')



#title at top
title = ttk.Label(st, text="Settings")
title.grid(column=0, row=0, padx=10, pady=10, sticky='w')



#excel file section
ef = ttk.LabelFrame(st, text='Excel File') #excel file frame
ef.grid(column=0, row=1, padx=10, pady=10, sticky='w')

ef_entry = ttk.Entry(ef, textvariable=ef_var, width=49, takefocus=False) #excel file entry input
ef_entry.grid(column=0, row=0, padx=10, pady=10)
ef_entry.bind("<FocusOut>", lambda event: config_instance.save('excel_file', ef_var.get(), True))
ef_entry.bind('<Control-a>', lambda x: ef_entry.selection_range(0, 'end') or "break")

ef_browse = ttk.Button(ef, text='Browse', command=select_excel_file, takefocus=False) #excel file browse button
ef_browse.grid(column=1, row=0, padx=10, pady=10)


#forms link section
fl = ttk.LabelFrame(st, text='G Forms Prefill link') #forms link frame
fl.grid(column=0, row=2, padx=10, pady=10, sticky='w')

fl_entry = ttk.Entry(fl, textvariable=fl_var, width=49, takefocus=False) #forms link entry input
fl_entry.bind("<FocusOut>", lambda event: config_instance.save('forms_link', fl_var.get(), True))
fl_entry.bind('<Control-a>', lambda x: fl_entry.selection_range(0, 'end') or "break")
fl_entry.grid(column=0, row=0, padx=10, pady=10)



#options section
op = ttk.Frame(st)
op.grid(column=0, row=3, padx=10, pady=10, sticky='w')


sc = ttk.Frame(op)#starting cell
sc.grid(column=0, row=0,padx=10, pady=10, sticky='w')
sc.grid_columnconfigure(0, weight=1)

sc_entry = ttk.Entry(sc, textvariable=sc_var, width=3, takefocus=False) #starting cell spinbox
sc_entry.bind("<FocusOut>", lambda event: config_instance.save('starting_cell', sc_var.get(), True))
sc_entry.grid(column=0, row=0)

sc_text = ttk.Label(sc, text='Starting Cell') #starting cell label
sc_text.grid(column=1, row=0)


bs = ttk.Frame(op)#box size frame
bs.grid(column=1, row=0,padx=10, pady=10, sticky='w')
bs.grid_columnconfigure(0, weight=1)

bs_spinbox = ttk.Spinbox(bs, textvariable=bs_var, from_=0, to=100, width=3, takefocus=False) #box size spinbox
bs_spinbox.bind("<FocusOut>", lambda event: config_instance.save('box_size', bs_var.get(), True))
bs_spinbox.grid(column=0, row=0)

bs_text = ttk.Label(bs, text='Box Size') #box size label
bs_text.grid(column=1, row=0)


br = ttk.Frame(op) #border size frame
br.grid(column=2, row=0,padx=10, pady=10, sticky='w')
br.grid_columnconfigure(0, weight=1)

br_spinbox = ttk.Spinbox(br, textvariable=br_var, from_=0, to=100, width=3, takefocus=False) #border size spinbox
br_spinbox.bind("<FocusOut>", lambda event: config_instance.save('border_size', br_var.get(), True))
br_spinbox.grid(column=0, row=0)

br_text = ttk.Label(br, text='Border Size') #border size label
br_text.grid(column=1, row=0)


#invert color checkbox
ic = ttk.Checkbutton(op,
                text='Invert Color',
                command=lambda: config_instance.save('invert_color', ic_var.get(), True),
                variable=ic_var,
                onvalue=True,
                offvalue=False,
                takefocus=False)
ic.grid(column=3, row=0, padx=10, pady=10)



#destination folder section
df = ttk.LabelFrame(st, text='Destination Folder') #df frame
df.grid(column=0, row=4, padx=10, pady=10, sticky='w')

df_entry = ttk.Entry(df, textvariable=df_var, width=49, takefocus=False) #df entry input
df_entry.grid(column=0, row=0, padx=10, pady=10)
df_entry.bind("<FocusOut>", lambda event: config_instance.save('destination', df_var.get(), True))
df_entry.bind('<Control-a>', lambda x: df_entry.selection_range(0, 'end') or "break")

df_browse = ttk.Button(df, text='Browse', command=select_destination, takefocus=False) #df browse button
df_browse.grid(column=1, row=0, padx=10, pady=10)




bt_label = ttk.Label(bt, text='Bitly Tokens')
bt_label.pack(side='top', anchor='nw', padx=10, pady=10)



ba = ttk.Frame(bt)
ba.pack(padx=10, pady=10)

ba_entry = ttk.Entry(ba, textvariable=bt_add, width=50, takefocus=False)
ba_entry.grid(column=0, row=0, padx=10)
ba_entry.bind('<Control-a>', lambda x: ba_entry.selection_range(0, 'end') or "break")
ba_button = ttk.Button(ba, text='Add', command=add_token, takefocus=False)
ba_button.grid(column=1, row=0, padx=10)


columns = ('tokens')
bt_tree = ttk.Treeview(bt, columns=columns, show=["headings"])
bt_tree.heading('tokens', text='Tokens')
#bt_tree.bind('<<bt_TreeviewSelect>>', item_selected)
bt_tree.pack(fill='x', expand=True, padx=20, pady=10)
show_token()

ll = ttk.Frame(bt)
ll.pack(fill='x', expand=True)

ll_left = ttk.Frame(ll)
ll_left.pack(side='left', padx=20)

usg = ttk.Label(ll_left)
usg.grid(sticky='w', pady=2)
usage = config_instance.get_usage()
usg.config(text=usage)


show_tokens = ttk.Checkbutton(ll_left,
                text='Show Tokens',
                command=show_token,
                variable=bt_show,
                onvalue=True,
                offvalue=False)
show_tokens.grid(row=1,sticky='w',pady=2)

bt_btn = ttk.Frame(ll)
bt_btn.pack(side='right', padx=10)

bt_btn_sl = ttk.Button(bt_btn, text='Select all', command=lambda: bt_tree.selection_set(tuple(bt_tree.get_children())), takefocus=False)
bt_btn_sl.grid(column=0, row=0, padx=10)

bt_btn_rm = ttk.Button(bt_btn, text='Remove', command=remove_token, takefocus=False)
bt_btn_rm.grid(column=1, row=0, padx=10)





cd_label = ttk.Label(cd, text='G Forms Codes')
cd_label.pack(side='top', anchor='nw', padx=10, pady=10)



ca = ttk.Frame(cd)
ca.pack(padx=10, pady=10)

ca_entry_cd = ttk.Entry(ca, textvariable=cd_add_cd, width=34, takefocus=False)
ca_entry_cd.grid(column=0, row=0, padx=5)
ca_entry_cd.bind('<Control-a>', lambda x: ca_entry_cd.selection_range(0, 'end') or "break")

ca_entry_rp = ttk.Entry(ca, textvariable=cd_add_rp, width=15, takefocus=False)
ca_entry_rp.grid(column=1, row=0, padx=5)
ca_entry_rp.bind('<Control-a>', lambda x: ca_entry_rp.selection_range(0, 'end') or "break")

ca_button = ttk.Button(ca, text='Add', command=add_code, takefocus=False)
ca_button.grid(column=2, row=0, padx=15)


columns = ('code', 'replacement')
cd_tree = ttk.Treeview(cd, columns=columns, show=["headings"])
cd_tree.heading('code', text='Code')
cd_tree.heading('replacement', text='Replacement Column')
#cd_tree.bind('<<cd_TreeviewSelect>>', item_selected)
cd_tree.pack(fill='x', expand=True, padx=20, pady=10)
show_code()

ll = ttk.Frame(cd)
ll.pack(fill='x', expand=True)

cd_btn = ttk.Frame(ll)
cd_btn.pack(side='right', padx=10)

cd_btn_sl = ttk.Button(cd_btn, text='Select all', command=lambda: cd_tree.selection_set(tuple(cd_tree.get_children())), takefocus=False)
cd_btn_sl.grid(column=0, row=0, padx=10)

cd_btn_rm = ttk.Button(cd_btn, text='Remove', command=remove_code, takefocus=False)
cd_btn_rm.grid(column=1, row=0, padx=10)





#generate button
gn = ttk.Button(root,
                text='Generate',
                command=generate,
                takefocus=False)
gn.pack(side='right', padx=30)



#progress section
ps = ttk.Frame(root) #progress section text
ps.pack(expand=True, side='bottom', fill='x')
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