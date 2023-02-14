from ast import For
from tkinter import *
import time
import os
root = Tk()

# smile frame : 25
# angry frame : 22, after=10


smile_frame_count = 25
angry_frame_count = 44

smile_frames = [PhotoImage(file='smile.gif',format = 'gif -index %i' %(i)) for i in range(smile_frame_count)]
angry_frames = [PhotoImage(file='angry.gif',format = 'gif -index %i' %(i)) for i in range(angry_frame_count)]

interval = 100
#interval = 10

def update(index):

	smile_frame = smile_frames[index]
	index += 1
	if index == smile_frame_count:
		index = 0
	label.configure(image=smile_frame)

	if index == smile_frame_count -1:
		i = 0
		while i < angry_frame_count:
			angry_frame = angry_frames[i]
			label.configure(image=angry_frame)
			label.update()
			i += 1
			time.sleep(0.02)

	root.after(interval, update, index)

label = Label(root)
label.pack()
root.after(0, update, 0)
root.title("귀찮다 v2.0")
root.attributes('-toolwindow', True)
root.mainloop()