from main import Config
from main import SimpleQR
from tkinter import ttk
import tkinter as tk
import threading
import webbrowser
import os

#root window
root = tk.Tk()
root.title('SimpleQR')
root.geometry('500x800+50+50')
root.minsize(500, 500)

config = Config('cfg.ini')
config.load()
link_var = tk.StringVar(root, config.url)
invert_var = tk.StringVar(root, config.invert)


def on_closing():
    config.save('url', link_var.get())
    config.save('invert', bool(int(invert_var.get())))
    root.destroy()


#title
title = ttk.Label(root, text="SimpleQR", font=("Arial", 25))
title.pack(fill='x', padx=5, pady=5)


#names
names_frame = ttk.LabelFrame(root, text='Names')
names_frame.pack(padx=5, pady=5, expand=True, fill='y')
names_text = tk.Text(names_frame, height=8, takefocus=False)
names_text.pack(padx=10, pady=(0, 10), expand=True, fill='y')


#forms links
fl = ttk.LabelFrame(root, text='G Forms Prefill link') #forms link frame
fl.pack(fill='x', padx=5, pady=5)

fl_entry = ttk.Entry(fl, textvariable=link_var, takefocus=False) #forms link entry input
fl_entry.bind("<FocusOut>", lambda event: config.save('url', link_var.get()))
fl_entry.bind('<Control-a>', lambda x: fl_entry.selection_range(0, 'end') or "break")
fl_entry.pack(padx=10, pady=10, expand=True, fill='x')


#invert color checkbox
ic = ttk.Checkbutton(root,
                text='Invert Color',
                command=lambda: config.save('invert', bool(int(invert_var.get()))),
                variable=invert_var,
                onvalue=1,
                offvalue=0,
                takefocus=False)
ic.pack(padx=10, pady=5, anchor='nw')


#generate sectopn
gs = ttk.Frame(root) #generate section frame
gs.pack(fill='x', anchor='s')

gn = ttk.Button(gs,
                text='Generate',
                command=print("GENERATE"),
                takefocus=False)
gn.pack(side='right', padx=20, pady=20)

ps = ttk.Frame(gs) #progress section text
ps.pack(expand=True, side='right', fill='x')
ps.columnconfigure(0, weight=1)

pt = ttk.Label(ps, text='') #progress text
pt.grid(column=0, row=0, padx=30, sticky='w')

pb = ttk.Progressbar( #progress bar
    ps,
    orient='horizontal',
    mode='determinate',
    length=480,
)


root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()