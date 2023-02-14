from tkinter import font
from tkinter import *



def shift_cursor1(event=None):
    position = entry.index(INSERT)
  
    # Changing position of cursor one character left
    entry.icursor(position+1)



root = Tk()

entry = Entry(root, font=("System", 11, 'bold'), fg="#707070", width=25, validatecommand=shift_cursor1)

entry.pack()
root.mainloop()
#print(font.families())
