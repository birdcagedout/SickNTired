from tkinter import * 
 
root = Tk()
root.geometry("350x200")
#frame = Frame(root)
#frame.pack()
 
#Scala = Scale(frame, from_=0, to=10)
#Scala.pack(padx=5, pady=5)

#entry = Entry(frame)
#entry.pack()

def updateScaler(event):
	scaleVal = scale.get()
	#entry.delete(0, 'end')
	#entry.insert(0, scaleVal)
	root.wm_attributes('-alpha', (100 - scaleVal)/100)
	root.update()

checkButtonVar = IntVar(value=1)
checkButton = Checkbutton(root, variable = checkButtonVar)
checkButton["state"] = DISABLED
checkButton.grid(row=0, column=0, padx=(35, 0))

label = Label(root, text="투명도 ")
label.grid(row=0, column=0, padx=(85,0))

scale = Scale(root, from_=0, to=90, resolution=10, showvalue=0, bd=1, orient=HORIZONTAL, length=200, #label="투명도(%)", 
				activebackground='#358CD6', troughcolor="#FDDB71")
scale.bind("<ButtonRelease-1>", updateScaler)
scale.grid(row=0, column=1)
 
root.mainloop()