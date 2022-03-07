import os
import sys
import time
import pyperclip
from tkinter import *
from pyautogui import *
from threading import *

WIN_WIDTH = 350
WIN_HEIGHT = 350

WAITING_WIN_WIDTH = 150
WAITING_WIN_HEIGHT = 150

BG_COLOR = 	'#dae5f7'	# '#D7ccf5'= 연한보라



class Bypass:
	def __init__(self):
		# tkinter
		self.root = Tk()
		self.root.iconbitmap(r'rocket.ico')
		self.root.geometry(f"{WIN_WIDTH}x{WIN_HEIGHT}+{int((self.root.winfo_screenwidth()-self.root.winfo_width())/2)}+{int((self.root.winfo_screenheight()-self.root.winfo_height())/2)}")
		self.root.title("귀찮다 v2.0")
		self.root.resizable(False, False)
		self.root.configure(bg=BG_COLOR)
		self.root.wm_attributes("-topmost", True)

		# "법인번호" 텍스트 라벨
		self.label_0 = Label(self.root, text="자동 입력할 조회 사유를 선택하세요", font=("Malgun Gothic", 15, 'bold'))
		self.label_0.configure(bg=BG_COLOR)

		self.radio_var = IntVar()
		self.reasons = ["", "신규등록 등 업무 대상자 확인", "이전등록 등 업무 대상자 확인", "말소압류 등 업무 대상자 확인", "보험 가입 여부 조회", ""]
		self.radio_btn1 = Radiobutton(self.root, text="신규등록 등 업무 대상자 확인", font=("Malgun Gothic", 12), value=1, variable=self.radio_var, command=self.on_select)
		self.radio_btn2 = Radiobutton(self.root, text="이전등록 등 업무 대상자 확인", font=("Malgun Gothic", 12), value=2, variable=self.radio_var, command=self.on_select)
		self.radio_btn3 = Radiobutton(self.root, text="말소압류 등 업무 대상자 확인", font=("Malgun Gothic", 12), value=3, variable=self.radio_var, command=self.on_select)
		self.radio_btn4 = Radiobutton(self.root, text="보험 가입 여부 조회", font=("Malgun Gothic", 12), value=4, variable=self.radio_var, command=self.on_select)
		self.radio_btn5 = Radiobutton(self.root, text="", font=("Malgun Gothic", 12), value=5, variable=self.radio_var, command=self.on_select)
		
		self.entry_5 = Entry(self.root, text="직접 입력 (한글 잘 안 됨~ㅋ)", font=("Malgun Gothic", 12), fg="#707070", width=25)
		self.entry_5.insert(0, "직접 입력 (한글 잘 안 됨~ㅋ)")

		#self.radio_btn1.bind("<Return>", self.on_start)
		#self.radio_btn2.bind("<Return>", self.on_start)
		#self.radio_btn3.bind("<Return>", self.on_start)
		#self.radio_btn4.bind("<Return>", self.on_start)
		#self.entry_5.bind("<Return>", self.on_start)
		self.entry_5.bind("<Button-1>", self.on_input)

		self.radio_btn1.configure(bg=BG_COLOR)
		self.radio_btn2.configure(bg=BG_COLOR)
		self.radio_btn3.configure(bg=BG_COLOR)
		self.radio_btn4.configure(bg=BG_COLOR)
		self.radio_btn5.configure(bg=BG_COLOR)

		self.label_0.grid(row=0, column=0, sticky=EW, padx=8, pady=5, columnspan=1)
		self.radio_btn1.grid(row=1, column=0, sticky=W, padx=35, pady=0)
		self.radio_btn2.grid(row=2, column=0, sticky=W, padx=35, pady=0)
		self.radio_btn3.grid(row=3, column=0, sticky=W, padx=35, pady=0)
		self.radio_btn4.grid(row=4, column=0, sticky=W, padx=35, pady=0)
		self.radio_btn5.grid(row=5, column=0, sticky=W, padx=35, pady=0)
		self.entry_5.grid(row=5, column=0, sticky=W, padx=62, pady=0)

		# 버튼
		self.btn_img = PhotoImage(file=r"btn1.png")
		self.start_btn = Button(self.root, image=self.btn_img, borderwidth=1)
		self.start_btn.bind("<ButtonRelease-1>", self.on_start)
		self.start_btn.configure(width=300, height=99, activebackground="#FFFFFF")
		self.start_btn.place(x=24, y=230)

		#self.start_btn.grid(row=6, column=0, sticky=EW, padx=0, pady=0)
		
		self.root.mainloop()

	def on_select(self):
		select = self.radio_var.get()
		
		if select == 5:
			self.entry_5.delete(0, END)
			self.entry_5.focus()
		else:
			self.entry_5.delete(0, END)
			self.entry_5.insert(0, "직접 입력 (한글 잘 안 됨~ㅋ)")
			self.root.focus()

	def on_input(self, event):
		self.entry_5.delete(0, END)
		self.radio_var.set(5)

	def on_start(self, event):
		if self.start_btn["state"] == DISABLED:
			return
		else:
			self.start_btn["state"] = DISABLED
			self.radio_btn1["state"] = DISABLED
			self.radio_btn2["state"] = DISABLED
			self.radio_btn3["state"] = DISABLED
			self.radio_btn4["state"] = DISABLED
			self.radio_btn5["state"] = DISABLED
			self.entry_5["state"] = DISABLED

		reason = ""
		choice = self.radio_var.get()
		
		if choice <= 4:
			reason = self.reasons[choice]
		else:
			reason = self.entry_5.get().strip()
		
		watcher = WatcherThread(reason=reason)
		watcher.start()

		self.win_move_aside()

	# 윈도 오른쪽 구석으로 이동
	def win_move_aside(self):
		pos_x = self.root.winfo_x()
		pos_y = self.root.winfo_y()

		scr_width = self.root.winfo_screenwidth()
		scr_height = self.root.winfo_screenheight()

		distance_x = scr_width - (pos_x + WIN_WIDTH) - 50 		# X거리(여유분=50)
		distance_y = scr_height - (pos_y + WIN_HEIGHT) - 100  	# Y거리(여유분=100)
		distance_left_x = distance_x
		distance_left_y = distance_y

		while distance_left_x > 0:
			self.root.geometry(f"+{pos_x + (distance_x - distance_left_x)}+{pos_y}")
			self.root.update()
			distance_left_x -= 2
		pos_x = self.root.winfo_x()

		while distance_left_y > 0:
			self.root.geometry(f"+{pos_x}+{pos_y + (distance_y - distance_left_y)}")
			self.root.update()
			distance_left_y -= 2

class WatcherThread(Thread):
	# 초기화
	def __init__(self, reason=""):
		super(WatcherThread, self).__init__(daemon=True)
		self.reason = reason

	def run(self):
		
		while True:
			# 둘 다 뜨는지 먼저 체크
			pos1 = locateCenterOnScreen("inquiry_popup.PNG", grayscale=True, region=(300, 300, 1200, 400))
			#pos2 = locateCenterOnScreen("inquiry_save.PNG", grayscale=True)
			if pos1 == None:
				time.sleep(0.3)
				continue

			# 조회 사유 자동입력
			pos2 = Point(x=pos1.x+410, y=pos1.y+40)
			#pos2 = Point(x=pos1.x+510, y=pos1.y+50)	# 강선임 주임님
			pos1 = Point(x=pos1.x, y=pos1.y+100)
			moveTo(pos1)
			click()

			# 현재 클립보드에 있던 내용을 미리 다른 곳에 저장해둠
			currentClipboard = pyperclip.paste()
			pyperclip.copy(self.reason)
			hotkey('ctrl', 'v')
			# 저장해둔 클립보드 내용을 현재 클립보드에 원래대로 복사해놓음
			pyperclip.copy(currentClipboard)
			#now = dt.fromtimestamp(time.time()).strftime('%Y-%m-%d %H:%M:%S')
			#print(f"[{now}] 사유 자동 입력 완료!")
			
			# 저장 클릭
			moveTo(pos2)
			click()
			#now = dt.fromtimestamp(time.time()).strftime('%Y-%m-%d %H:%M:%S')
			#print(f"[{now}] 저장 버튼 클릭 완료!")
			
			# 2초 쉬기
			print("")
			time.sleep(2)


if __name__ == "__main__":
	#try:
		#os.chdir(sys._MEIPASS)
		#print(sys._MEIPASS)
	#except:
		#os.chdir(os.getcwd())
		#print(os.getcwd())
	app = Bypass()
