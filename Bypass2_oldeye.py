import os
import sys
import time
import pyperclip
from tkinter import *
from pyautogui import *
from threading import *

WIN_WIDTH = 350
WIN_HEIGHT = 350

WAITING_WIN_WIDTH = 200
WAITING_WIN_HEIGHT = 200

BG_COLOR = 	'#dae5f7'	# '#D7ccf5'= 연한보라

ACTIVATED = False

HOME_PATH = os.path.dirname(__file__)
VER = "2.1"


class Bypass:
	def __init__(self):
		# tkinter
		self.root = Tk()
		self.root.iconbitmap(os.path.join(HOME_PATH, "tube.ico"))
		self.root.geometry(f"{WIN_WIDTH}x{WIN_HEIGHT}+{int((self.root.winfo_screenwidth()-self.root.winfo_width())/2)}+{int((self.root.winfo_screenheight()-self.root.winfo_height())/2)}")
		self.root.title(f"귀찮다 v{VER}")
		self.root.resizable(False, False)
		self.root.configure(bg=BG_COLOR)
		self.root.wm_attributes("-topmost", True)

		# "법인번호" 텍스트 라벨
		self.label_0 = Label(self.root, text="자동 입력할 조회 사유를 선택하세요", font=("Malgun Gothic", 12, 'bold'))
		self.label_0.configure(bg=BG_COLOR)

		self.radio_var = IntVar()
		self.reasons = ["", "신규등록 등 업무 대상자 확인", "이전등록 등 업무 대상자 확인", "말소압류 등 업무 대상자 확인", "보험 가입 여부 조회", ""]
		self.radio_btn1 = Radiobutton(self.root, text="신규등록 등 업무 대상자 확인", font=("Malgun Gothic", 10), value=1, variable=self.radio_var, command=self.on_select)
		self.radio_btn2 = Radiobutton(self.root, text="이전등록 등 업무 대상자 확인", font=("Malgun Gothic", 10), value=2, variable=self.radio_var, command=self.on_select)
		self.radio_btn3 = Radiobutton(self.root, text="말소압류 등 업무 대상자 확인", font=("Malgun Gothic", 10), value=3, variable=self.radio_var, command=self.on_select)
		self.radio_btn4 = Radiobutton(self.root, text="보험 가입 여부 조회", font=("Malgun Gothic", 10), value=4, variable=self.radio_var, command=self.on_select)
		self.radio_btn5 = Radiobutton(self.root, text="", font=("Malgun Gothic", 10), value=5, variable=self.radio_var, command=self.on_select)
		
		self.entry_5 = Entry(self.root, font=("Malgun Gothic", 10), fg="#707070", width=25)
		self.entry_5.insert(0, "직접 입력(한글 잘 안 되지롱~)")
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
		self.btn_img = PhotoImage(file=os.path.join(HOME_PATH, "btn1.png"))
		self.start_btn = Button(self.root, image=self.btn_img, borderwidth=1)
		self.start_btn.bind("<ButtonRelease-1>", self.on_start)
		self.start_btn.configure(width=300, height=99, activebackground="#FFFFFF")
		#self.start_btn.place(x=24, y=230)
		self.start_btn.grid(row=6, column=0, sticky=EW, padx=25, pady=20, columnspan=1)

		# 대기 화면
		self.canvas = Canvas(self.root, width=WAITING_WIN_HEIGHT, height=WAITING_WIN_HEIGHT, bg='white', bd=0, highlightthickness=0)
	
		self.smile_frame_count = 25
		self.angry_frame_count = 44

		self.smile_frames = [PhotoImage(file=os.path.join(HOME_PATH, "smile.gif"), format = 'gif -index %i' %(i)) for i in range(self.smile_frame_count)]
		self.angry_frames = [PhotoImage(file=os.path.join(HOME_PATH, "angry.gif"), format = 'gif -index %i' %(i)) for i in range(self.angry_frame_count)]

		#self.label_1 = Label(self.canvas)
		#self.label_1.configure(bg="white")
		
		#self.label_1.lower()
		#self.root.update()
		
		self.root.mainloop()

	def on_select(self):
		select = self.radio_var.get()
		
		if select == 5:
			self.entry_5.delete(0, END)
			self.entry_5.focus()
		else:
			self.entry_5.delete(0, END)
			self.entry_5.insert(0, "직접 입력(한글 잘 안 되지롱~)")
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
		self.wait()

	# 윈도 오른쪽 구석으로 이동
	def win_move_aside(self):
		pos_x = self.root.winfo_x()
		pos_y = self.root.winfo_y()

		scr_width = self.root.winfo_screenwidth()
		scr_height = self.root.winfo_screenheight()

		distance_x = scr_width - (pos_x + WIN_WIDTH) + 100		# X거리(여유분=50)
		distance_y = scr_height - (pos_y + WIN_HEIGHT) + 50  	# Y거리(여유분=50)
		distance_left_x = distance_x
		distance_left_y = distance_y

		while distance_left_x > 0:
			self.root.geometry(f"+{pos_x + (distance_x - distance_left_x)}+{pos_y}")
			self.root.update()
			distance_left_x -= 5
		pos_x = self.root.winfo_x()

		while distance_left_y > 0:
			self.root.geometry(f"+{pos_x}+{pos_y + (distance_y - distance_left_y)}")
			self.root.update()
			distance_left_y -= 5
		
	def wait(self):
		self.label_0.destroy()
		self.radio_btn1.destroy()
		self.radio_btn2.destroy()
		self.radio_btn3.destroy()
		self.radio_btn4.destroy()
		self.radio_btn5.destroy()
		self.entry_5.destroy()
		self.start_btn.destroy()
		self.root.update()


		while self.root.winfo_width() > WAITING_WIN_WIDTH:
			self.root.geometry(f"{self.root.winfo_width() - 5}x{self.root.winfo_height() - 5}")
			self.root.update()
		self.root.title(f"귀찮다 v{VER} 작동 중")
		self.root.attributes('-toolwindow', True)


		image_on_canvas = self.canvas.create_image(0, 0, anchor=NW, image=self.smile_frames[0])
		self.canvas.pack(fill="both", expand=True)
		self.root.update()

		i = 0
		while True:
			smile_frame = self.smile_frames[i]
			self.canvas.itemconfigure(image_on_canvas, image=smile_frame)
			self.root.update()

			i += 1
			if i == self.smile_frame_count:
				i = 0

			global ACTIVATED
			if ACTIVATED == True:
				ACTIVATED = False

				# 창 맨 위로
				self.root.wm_attributes("-topmost", True)
				self.root.update()

				j = 0
				while j < self.angry_frame_count:
					angry_frame = self.angry_frames[j]
					self.canvas.itemconfigure(image_on_canvas, image=angry_frame)
					self.root.update()
					j += 1
					time.sleep(0.02)
			
			time.sleep(0.1)


class WatcherThread(Thread):
	# 초기화
	def __init__(self, reason=""):
		super(WatcherThread, self).__init__(daemon=True)
		self.reason = reason

	def run(self):
		
		while True:
			# 둘 다 뜨는지 먼저 체크
			pos1 = locateCenterOnScreen(os.path.join(HOME_PATH,"inquiry_popup_oldeye.PNG"), grayscale=True, region=(600, 300, 720, 480))
			#pos2 = locateCenterOnScreen("inquiry_save.PNG", grayscale=True)
			if pos1 == None:
				time.sleep(0.3)
				continue

			# 작동중 ==> 메인 클래스에 알림 
			global ACTIVATED
			ACTIVATED = True
			
			# 조회 사유 자동입력
			#pos2 = Point(x=pos1.x+410, y=pos1.y+40)
			pos2 = Point(x=pos1.x+510, y=pos1.y+50)	# 강선임 주임님
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
			#print("")
			time.sleep(2)


if __name__ == "__main__":
	#try:
		#os.chdir(sys._MEIPASS)
		#print(sys._MEIPASS)
	#except:
		#os.chdir(os.getcwd())
		#print(os.getcwd())
	app = Bypass()
