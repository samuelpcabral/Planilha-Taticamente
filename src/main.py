import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
import os


# create the root window
root = tk.Tk()
root.title('Selecione os arquivos html')
root.resizable(False, False)
root.geometry('300x150')

files = None

def select_files():
    filetypes = (
        ('html files', '*.html'),
        ('All files', '*.*')
    )

    filenames = fd.askopenfilenames(
        title='Open files',
        initialdir=os.getcwd(),
        filetypes=filetypes)

    files = filenames
    # showinfo(
    #     title='Selected Files',
    #     message=filenames
    # )


# open button
open_button = ttk.Button(
    root,
    text='Open Files',
    command=select_files
)
print(files)

open_button.pack(expand=True)

root.mainloop()

import tkinter as tk
from tkinter import filedialog

root= tk.Tk()

def openFile():
    files = filedialog.askopenfilenames(parent=root, title='Choose a file')
    print (files)

button = tk.Button(root, text="Open File", command=openFile)
button.pack()

root.mainloop()