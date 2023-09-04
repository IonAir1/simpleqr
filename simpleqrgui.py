from simpleqr import Config
from simpleqr import SimpleQR
from tkinter import ttk
import tkinter as tk
import threading
import webbrowser
import os

#root window
root = tk.Tk()
root.title('SimpleQR')
root.geometry('500x600+50+50')
root.minsize(500, 500)

config = Config('cfg.ini')
config.load()
link_var = tk.StringVar(root, config.url)
invert_var = tk.StringVar(root, config.invert)
split_var = tk.StringVar(root, config.split)


#Extended Text Widget that includes a context menu and select all support
class TextX(tk.Text):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.menu = tk.Menu(self, tearoff=False)
        self.menu.add_command(label="Undo", command=self.popup_undo)
        self.menu.add_command(label="Redo", command=self.popup_redo)
        self.menu.add_separator()
        self.menu.add_command(label="Select All", command=lambda: self.select_all(''))
        self.menu.add_command(label="Cut", command=self.popup_cut)
        self.menu.add_command(label="Copy", command=self.popup_copy)
        self.menu.add_command(label="Paste", command=self.popup_paste)
        self.menu.bind("<FocusOut>",lambda x: self.menu.unpost())
        self.bind("<Button-3>", self.display_popup)
        self.bind("<Control-Key-a>", self.select_all)
        self.bind("<Control-Key-A>", self.select_all)
        self.bind('<Control-v>', self.paste)
        self.bind('<Control-V>', self.paste)

    def select_all(self, event):
        self.tag_add(tk.SEL, "1.0", tk.END)
        self.mark_set(tk.INSERT, "1.0")
        self.see(tk.INSERT)
        return 'break'

    def display_popup(self, event):
        self.menu.post(event.x_root, event.y_root)
        self.menu.focus_set()

    def popup_undo(self):
        self.event_generate("<<Undo>>")
        self.menu.unpost()

    def popup_redo(self):
        self.event_generate("<<Redo>>")
        self.menu.unpost()

    def popup_copy(self):
        self.event_generate("<<Copy>>")
        self.menu.unpost()

    def popup_cut(self):
        self.event_generate("<<Cut>>")
        self.menu.unpost()

    def popup_paste(self):
        self.event_generate("<<Paste>>")
        self.paste(False)
        self.menu.unpost()

    def paste(self, e):
        if self.tag_ranges(tk.SEL):
            self.delete(tk.SEL_FIRST, tk.SEL_LAST)


#Extended Text Widget that includes a context menu and select all support
class EntryX(tk.Entry):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.menu = tk.Menu(self, tearoff=False)
        self.menu.add_command(label="Select All", command=lambda: self.select_range(0, tk.END))
        self.menu.add_command(label="Cut", command=self.popup_cut)
        self.menu.add_command(label="Copy", command=self.popup_copy)
        self.menu.add_command(label="Paste", command=self.popup_paste)
        self.menu.bind("<FocusOut>",lambda x: self.menu.unpost())
        self.bind("<Button-3>", self.display_popup)
        self.bind("<Control-Key-a>", lambda: self.select_range(0, tk.END))
        self.bind("<Control-Key-A>", lambda: self.select_range(0, tk.END))
        self.bind('<Control-v>', self.paste)
        self.bind('<Control-V>', self.paste)

    def display_popup(self, event):
        self.menu.post(event.x_root, event.y_root)
        self.menu.focus_set()

    def popup_copy(self):
        self.event_generate("<<Copy>>")
        self.menu.unpost()

    def popup_cut(self):
        self.event_generate("<<Cut>>")
        self.menu.unpost()

    def popup_paste(self):
        self.event_generate("<<Paste>>")
        self.paste(False)
        self.menu.unpost()

    def paste(self, e):
        if self.tag_ranges(tk.SEL):
            self.delete(tk.SEL_FIRST, tk.SEL_LAST)


def on_closing():
    config.save('url', link_var.get())
    config.save('invert', bool(int(invert_var.get())))
    root.destroy()


def generate():
    names = names_text.get('1.0', tk.END)
    simpleqr = SimpleQR(names, link_var.get())
    generate_thread = threading.Thread(target=generate_thread_function, kwargs={'instance':simpleqr, 'names':names})
    generate_thread.start()

    progress_thread = threading.Thread(target=progress_function, kwargs={'instance':simpleqr, 'thread':generate_thread})
    progress_thread.start()


def progress_function(**kwargs):
    simpleqr = kwargs.get('instance')
    generate_thread = kwargs.get('thread')

    while generate_thread.is_alive():
        # print(simpleqr.progress)
        pb['value'] = simpleqr.progress

def generate_thread_function(**kwargs):
    simpleqr = kwargs.get('instance')
    names = kwargs.get('names')
    if names.startswith("https://"):
        simpleqr.generate(invert=invert_var.get(), split=split_var.get(), replace=False)
    else:
        simpleqr.generate(invert=invert_var.get(), split=split_var.get())
    
    webbrowser.open('file://' + os.getcwd().replace('\\', '/') + '/exports')


#title
title = ttk.Label(root, text="SimpleQR", font=("Arial", 25))
title.pack(fill='x', padx=5, pady=5)


#names
names_frame = ttk.LabelFrame(root, text='Names')
names_frame.pack(padx=5, pady=5, expand=True, fill='y')
names_text = TextX(names_frame, height=8, takefocus=False, undo=True)
names_text.pack(padx=10, pady=(0, 10), expand=True, fill='y')


#forms links
fl = ttk.LabelFrame(root, text='G Forms Prefill link') #forms link frame
fl.pack(fill='x', padx=5, pady=5)

fl_entry = EntryX(fl, textvariable=link_var, takefocus=False) #forms link entry input
fl_entry.bind("<FocusOut>", lambda event: config.save('url', link_var.get()))
fl_entry.bind('<Control-a>', lambda x: fl_entry.selection_range(0, 'end') or "break")
fl_entry.bind('<Control-A>', lambda x: fl_entry.selection_range(0, 'end') or "break")
fl_entry.pack(padx=10, pady=10, expand=True, fill='x')

buttons = ttk.Frame()
buttons.pack(anchor='nw')

#invert color checkbox
ic = ttk.Checkbutton(buttons,
                text='Invert Color',
                command=lambda: config.save('invert', bool(int(invert_var.get()))),
                variable=invert_var,
                onvalue=1,
                offvalue=0,
                takefocus=False)
ic.grid(row=0, column=0, padx=10, pady=5)



#split name checkbox
sn = ttk.Checkbutton(buttons,
                text='Split Name',
                command=lambda: config.save('split', bool(int(split_var.get()))),
                variable=split_var,
                onvalue=1,
                offvalue=0,
                takefocus=False)
sn.grid(row=0, column=1, padx=10, pady=5)


#generate sectopn
gs = ttk.Frame(root) #generate section frame
gs.pack(fill='x', anchor='s')

gn = ttk.Button(gs,
                text='Generate',
                command=lambda: generate(),
                takefocus=False)
gn.pack(side='right', padx=(5, 10), pady=20)


pb = ttk.Progressbar( #progress bar
    gs,
    orient='horizontal',
    mode='determinate',
    length=480,
)
pb.pack(side='right', expand=True, fill='x', padx=(10,5))


root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()