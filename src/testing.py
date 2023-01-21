from tkinter import *

# Create the root window
root = Tk()

# Title of the window
root.title("Nice GUI")

# Size of the window
root.geometry("400x400")

# Create a text label
label = Label(root, text="This is a nice GUI!")

# Create a button
button = Button(root, text="Click Me!")

# Place the label and button on the screen
label.pack()
button.pack()

# Keep the window open
root.mainloop()