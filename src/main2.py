import os
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.title("Select HTML Files")
root.geometry("300x150")

def openFile():
    files = filedialog.askopenfilenames(parent=root, title='Choose a file', filetypes=[("HTML Files", "*.html")])
    print(files)

button = tk.Button(root, text="Open File", command=openFile)
button.pack()

root.mainloop()