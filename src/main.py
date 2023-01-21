import os
import tkinter as tk
from tkinter import filedialog, Label
import FillExcel

root = tk.Tk()
root.title("Selecione os arquivos HTML")
root.geometry("300x150")
root.resizable(False, False)

label = Label(root, text="TESTE TESTE TESTE")
label.pack(anchor='n')


def open_file():
    files = filedialog.askopenfilenames(parent=root, title='Choose a file', filetypes=[("HTML Files", "*.html")],
                                        initialdir=os.getcwd())
    FillExcel.create_sheets(files)
    label2 = Label(root, text="teste arquivos excel")
    label2.place(relx=0.5, rely=0.9, anchor='s')


button = tk.Button(root, text="Selecione os arquivos", command=open_file, font=("Inter", 12))
button.place(relx=0.5, rely=0.5, anchor='center')

root.mainloop()
