import os
import sys
import time
import traceback
import keyboard
import mouse
import ctypes
from tkinter import *
from tkinter import messagebox
from threading import *
import win32gui, win32con

WIN_WIDTH = 350
WIN_HEIGHT = 350

WAITING_WIN_WIDTH = 200
WAITING_WIN_HEIGHT = 200

BG_COLOR = 	'#f5fbff'	# '#D7ccf5'= 연한보라

ACTIVATED = False
#THREAD_EXPIRED = False
#THREAD_DEAD = False

SCREEN_WIDTH = 1920
SCREEN_HEIGHT = 1080
#REGION = (230, 230, 1200, 400)

HOME_PATH = os.path.dirname(__file__)
VER = "3.0"								# Pyautogui 제거, keyboard 모듈 추가, mouse 모듈 추가

# 인트로 스플래시 이미지 닫기(python3.9부터 importlib --> importlib.util로 바뀜)
import importlib.util
if '_PYIBoot_SPLASH' in os.environ and importlib.util.find_spec("pyi_splash"):
	import pyi_splash
	#pyi_splash.update_text('UI Loaded ...')
	pyi_splash.close()


# prevent the computer from going to sleep
#ctypes.windll.kernel32.SetThreadExecutionState(0x80000003)


# 디스플레이 DPI 조정 보정 ==> 해상도 다른 컴터에서 안 먹힘
#ctypes.windll.shcore.SetProcessDpiAwareness(2)


class Bypass:
	def __init__(self):
		# tkinter
		self.root = Tk()
		self.root.iconbitmap(os.path.join(HOME_PATH, "tube.ico"))
		self.root.tk.call('tk', 'scaling', 96/72)					# 해상도와 무관하게 DPI scaling = 96/72
		self.root.geometry(f"{WIN_WIDTH}x{WIN_HEIGHT}+{int((self.root.winfo_screenwidth()-WIN_WIDTH)/2-3)}+{int((self.root.winfo_screenheight()-WIN_HEIGHT)/2-3)}")
		self.root.title(f"귀찮다 v{VER}")
		self.root.resizable(False, False)
		self.root.configure(bg=BG_COLOR)
		self.root.wm_attributes("-topmost", True)

		# "사유를 선택하세요" 텍스트 라벨
		self.label_0 = Label(self.root, text="자동 입력할 조회 사유를 선택하세요", font=("Malgun Gothic", 15, 'bold'))

		self.radio_var = IntVar()
		self.reasons = ["", "신규등록 등 업무 대상자 확인", "이전등록 등 업무 대상자 확인", "보험 가입 여부 확인", "소유자 확인"]	
		self.radio_btn1 = Radiobutton(self.root, text="신규등록 등 업무 대상자 확인", font=("Malgun Gothic", 12), value=1, variable=self.radio_var, command=self.on_select)
		self.radio_btn2 = Radiobutton(self.root, text="이전등록 등 업무 대상자 확인", font=("Malgun Gothic", 12), value=2, variable=self.radio_var, command=self.on_select)
		self.radio_btn3 = Radiobutton(self.root, text="보험 가입 여부 확인", font=("Malgun Gothic", 12), value=3, variable=self.radio_var, command=self.on_select)
		self.radio_btn4 = Radiobutton(self.root, text="소유자 확인", font=("Malgun Gothic", 12), value=4, variable=self.radio_var, command=self.on_select)
		#self.radio_btn5 = Radiobutton(self.root, text="", font=("Malgun Gothic", 12), value=5, variable=self.radio_var, command=self.on_select)

		# 투명도 콘트롤: 체크버튼 + 라벨 + 스케일
		self.checkButtonVar = IntVar(value=1)
		self.checkButton = Checkbutton(self.root, variable = self.checkButtonVar, bg=BG_COLOR)
		self.checkButton["state"] = DISABLED

		self.label_1 = Label(self.root, text="투명도 ", font=("Malgun Gothic", 11, 'bold'), fg="#FDAA00")
		self.scale = Scale(self.root, from_=0, to=90, resolution=10, showvalue=0, bd=2, orient=HORIZONTAL, length=180,
							activebackground='#358CD6', highlightbackground=BG_COLOR, troughcolor="#FDDB71", command=self.updateScale)
		
		#self.entry_5 = Entry(self.root, font=("Malgun Gothic", 12), fg="#707070", width=25)
		#self.entry_5.insert(0, "직접 입력(한글 잘 안 되지롱~)")
		#self.entry_5.bind("<Button-1>", self.on_input)

		# 시작 버튼
		#self.btn_img = PhotoImage(file=os.path.join(HOME_PATH, "btn1.png"))
		#self.start_btn = Button(self.root, image=self.btn_img, borderwidth=1)
		self.start_btn = Button(self.root, text="고고씽ㅋ~", font=("Malgun Gothic", 35, 'bold'), foreground="#FFFFFF", background="#208DDE", activeforeground="#DDDDDD", activebackground="#55b6fd", borderwidth=1)
		self.start_btn.bind("<ButtonRelease-1>", self.on_start)

		# 개발자 로고
		self.label_2 = Label(self.root, text="developed by 김재형, 2022", font=("Malgun Gothic", 8))

		# 모든 위젯(콘트롤)의 배경색 통일
		self.label_0.configure(bg=BG_COLOR)
		self.radio_btn1.configure(bg=BG_COLOR)
		self.radio_btn2.configure(bg=BG_COLOR)
		self.radio_btn3.configure(bg=BG_COLOR)
		self.radio_btn4.configure(bg=BG_COLOR)
		#self.radio_btn5.configure(bg=BG_COLOR)
		self.label_1.configure(bg=BG_COLOR)
		self.scale.configure(bg=BG_COLOR)
		self.label_2.configure(bg=BG_COLOR)

		# 모든 위젯(콘트롤)의 layout 설정
		self.label_0.grid(row=0, column=0, sticky=EW, padx=5, pady=(5, 5))
		self.radio_btn1.grid(row=1, column=0, sticky=W, padx=35, pady=0)
		self.radio_btn2.grid(row=2, column=0, sticky=W, padx=35, pady=0)
		self.radio_btn3.grid(row=3, column=0, sticky=W, padx=35, pady=0)
		self.radio_btn4.grid(row=4, column=0, sticky=W, padx=35, pady=0)
		#self.radio_btn5.grid(row=5, column=0, sticky=W, padx=35, pady=0)
		#self.entry_5.grid(row=5, column=0, sticky=W, padx=62, pady=0)
		self.checkButton.grid(row=5, column=0, sticky=W, padx=(35, 0), pady=(1,5))
		self.label_1.grid(row=5, column=0, sticky=W, padx=(57,0), pady=(1,5))
		self.scale.grid(row=5, column=0, sticky=W, padx=(110, 0), pady=(1,5))
		self.start_btn.grid(row=6, column=0, sticky=EW, padx=(30, 30), pady=5)
		self.label_2.grid(row=7, column=0, sticky=E, padx=30, pady=0)

		# 대기 화면
		self.canvas = Canvas(self.root, width=WAITING_WIN_HEIGHT, height=WAITING_WIN_HEIGHT, bg='white', bd=0, highlightthickness=0)
	
		self.smile_frame_count = 25
		self.angry_frame_count = 44

		self.smile_frames = [PhotoImage(file=os.path.join(HOME_PATH, "smile.gif"), format = 'gif -index %i' %(i)) for i in range(self.smile_frame_count)]
		self.angry_frames = [PhotoImage(file=os.path.join(HOME_PATH, "angry.gif"), format = 'gif -index %i' %(i)) for i in range(self.angry_frame_count)]

		# 사유
		self.reason = ""
		self.watcher = None

		# 투명도
		self.alpha = 1		# 기본값 = 불투명

		# 자동차관리정보시스템 창 얻기
		self.target_win_hWnd = None
		self.target_win_title = None
		
		self.get_target_win()
		if self.target_win_hWnd == None:
			messagebox.showerror("확인", "자동차(이륜차)관리정보시스템이 시작되지 않았습니다.\n프로그램을 종료합니다.")
			sys.exit(0)
		
		# 최초 윈도 = 맨 앞에(뒤로 가는 경우도 있어서)
		self.root.lift()
		
		self.root.mainloop()


	# 투명도 스케일 변경
	def updateScale(self, event):
		self.scale.set(self.scale.get())
		self.alpha = self.scale.get()
		self.root.wm_attributes("-alpha", (100 - self.alpha)/100)
		self.root.update()


	# 윈도 스타일 바꾸기
	def setWinStyle(self, root):
		set_window_pos = ctypes.windll.user32.SetWindowPos
		set_window_long = ctypes.windll.user32.SetWindowLongPtrW
		get_window_long = ctypes.windll.user32.GetWindowLongPtrW
		get_parent = ctypes.windll.user32.GetParent

		hWnd = get_parent(root.winfo_id())

		old_style = get_window_long(hWnd, win32con.GWL_STYLE) # Get the style
		new_style = old_style & ~win32con.WS_MAXIMIZEBOX & ~win32con.WS_MINIMIZEBOX # New style, without max/min buttons

		set_window_long(hWnd, win32con.GWL_STYLE, new_style) # Apply the new style
		set_window_pos(hWnd, 0, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOZORDER | win32con.SWP_FRAMECHANGED)		# Updates


	# 현재 (보이는 윈도 중에서) 원하는 윈도 찾기 : "자동차관리정보시스템" 문자열 포함
	def get_target_win(self):
		def winEnumHandler(hWnd, ctx):
			if win32gui.IsWindowVisible(hWnd):
				title = win32gui.GetWindowText(hWnd)
				if "자동차관리정보시스템" in title or "이륜차관리정보시스템" in title:
					self.target_win_hWnd = hWnd
					self.target_win_title = title
		win32gui.EnumWindows(winEnumHandler, None)


	# 라디오버튼 클릭
	def on_select(self):
		select = self.radio_var.get()
		
		#if select == 5:
		#	self.entry_5.delete(0, END)
		#	self.entry_5.focus()
		#else:
		#	self.entry_5.delete(0, END)
		#	self.entry_5.insert(0, "직접 입력(한글 잘 안 되지롱~)")
		#	self.root.focus()
	

	# 직접 사유 입력
	def on_input(self, event):
		#self.entry_5.delete(0, END)
		#self.radio_var.set(5)
		pass
	

	# 시작버튼 클릭
	def on_start(self, event):

		# 라디오 버튼 선택은 1~4 (선택하지 않으면 0)
		choice = self.radio_var.get()

		if choice == 0:
			messagebox.showerror("확인", "자동 입력할 사유를 선택해주세요.")
			return
		else:
			self.reason = self.reasons[choice]
		#elif choice == 5:
		#	self.reason = self.entry_5.get().strip()
			

		#call to change style after the mainloop started. Directly call setWinStyle will not work.
		#self.root.after(10, lambda: self.setWinStyle(self.root)) 
		self.setWinStyle(self.root)

		if self.start_btn["state"] == DISABLED:
			return
		else:
			self.start_btn["state"] = DISABLED
			self.radio_btn1["state"] = DISABLED
			self.radio_btn2["state"] = DISABLED
			self.radio_btn3["state"] = DISABLED
			self.radio_btn4["state"] = DISABLED
			#self.radio_btn5["state"] = DISABLED
			#self.entry_5["state"] = DISABLED
			self.checkButton["state"] = DISABLED
			self.label_1["state"] = DISABLED
			self.scale["state"] = DISABLED
		
		self.win_move_aside()

		self.watcher = WatcherThread(reason=self.reason)
		self.watcher.start()

		self.wait()

	# 윈도 오른쪽 구석으로 이동
	def win_move_aside(self):
		pos_x = self.root.winfo_x()
		pos_y = self.root.winfo_y()

		scr_width = self.root.winfo_screenwidth()
		scr_height = self.root.winfo_screenheight()

		distance_x = scr_width - (pos_x + WAITING_WIN_WIDTH) - 10		# X거리(여유분=10)
		distance_y = scr_height - (pos_y + WAITING_WIN_HEIGHT) - 75  	# Y거리(여유분=100)
		distance_left_x = distance_x
		distance_left_y = distance_y

		while distance_left_x > 0:
			self.root.geometry(f"+{pos_x + (distance_x - distance_left_x)}+{pos_y}")
			self.root.update()
			distance_left_x -= 10
		pos_x = self.root.winfo_x()

		while distance_left_y > 0:
			self.root.geometry(f"+{pos_x}+{pos_y + (distance_y - distance_left_y)}")
			self.root.update()
			distance_left_y -= 10
		
	# 오리 춤추는 대기 화면
	def wait(self):
		self.label_0.destroy()
		self.radio_btn1.destroy()
		self.radio_btn2.destroy()
		self.radio_btn3.destroy()
		self.radio_btn4.destroy()
		#self.radio_btn5.destroy()
		#self.entry_5.destroy()
		self.checkButton.destroy()
		self.label_1.destroy()
		self.scale.destroy()
		self.start_btn.destroy()
		self.label_2.destroy()
		self.root.update()

		# 윈도 크기를 오리크기(WAITING_WIN_WIDTH)로 줄임
		while self.root.winfo_width() > WAITING_WIN_WIDTH:
			self.root.geometry(f"{self.root.winfo_width() - 5}x{self.root.winfo_height() - 5}")
			self.root.update()
		
		# 작은 윈도 모드 + 항상위 속성 제거
		self.root.title(f"귀찮다 v{VER} 작동 중")
		self.root.attributes("-topmost", False)

		# 오리 스마일 세팅
		i = 0
		image_on_canvas = self.canvas.create_image(0, 0, anchor=NW, image=self.smile_frames[i])
		self.canvas.grid(row=0, column=0, padx=0, pady=0, sticky=NW)
		self.root.update()

		while True:
			# 오리 스마일 시작
			smile_frame = self.smile_frames[i]

			# 윈도 종료시 TclError발생: _tkinter.TclError: invalid command name ".!canvas"
			try:
				self.canvas.itemconfigure(image_on_canvas, image=smile_frame)
			except TclError:
				return
			self.root.update()

			i += 1
			if i == self.smile_frame_count:
				i = 0
				#self.root.update()
				#ctypes.windll.kernel32.SetThreadExecutionState(0x80000003)

			# 팝업창 Title 발견!!
			global ACTIVATED
			if ACTIVATED == True:
				ACTIVATED = False

				# 윈도 맨 앞으로
				self.root.lift()
				self.root.attributes("-topmost", True)
				self.root.update_idletasks()
				#self.root.after_idle(self.root.attributes,"-topmost", False)
				self.root.after(0, self.root.attributes,"-topmost", False)

				# 오리 빡침 시작
				j = 0
				while j < self.angry_frame_count:
					angry_frame = self.angry_frames[j]
					self.canvas.itemconfigure(image_on_canvas, image=angry_frame)
					self.root.update()
					j += 1
					time.sleep(0.01)

			time.sleep(0.1)


class WatcherThread(Thread):
	# 초기화
	def __init__(self, reason=""):
		super(WatcherThread, self).__init__(daemon=True)
		self.reason = reason

	def run(self):

		# 이미 팝업창이 떠 있는 상태에서 실행된 경우 ==> 최상위 윈도가 현재 "귀찮다"이므로 팝업창 감지 안됨
		# 따라서 쓰레드 실행시 팝업창이 있다면 최상위로 올려주어 감지
		def winEnumHandler(hWnd, ctx):
			if win32gui.IsWindowVisible(hWnd):
				title = win32gui.GetWindowText(hWnd)
				if ("조회 사유 입력" in title) or ("개인정보 공개 열람사유입력" in title):
					win32gui.SetForegroundWindow(hWnd)
		win32gui.EnumWindows(winEnumHandler, None)

		while True:

			# 팝업창 Title 찾기
			try:
				hWnd = win32gui.GetForegroundWindow()
				title = win32gui.GetWindowText(hWnd)
				rect = win32gui.GetWindowRect(hWnd)
			except Exception:
				time.sleep(0.1)
				continue

			x = rect[0]
			y = rect[1]
			w = rect[2] - x
			h = rect[3] - y

			# 둘 다 아닌 경우
			if (title != "조회 사유 입력") and (title != "개인정보 공개 열람사유입력"):
				time.sleep(0.1)
				continue
			
			# 둘 중 하나는 발견한 경우: 너무 빠르니까 작동오류 발생 ==> 지연시간 추가
			time.sleep(0.2)
			
			# 작동중 ==> 메인 쓰레드에 알림 
			global ACTIVATED
			ACTIVATED = True
				
			# 사유 입력창 클릭
			ctypes.windll.user32.BlockInput(True)				# 마우스/키보드 입력 전면 차단
			xPos, yPos = mouse.get_position()
			mouse.move(x + 110, y + 110)
			mouse.click()
			keyboard.write(self.reason)
			time.sleep(0.1)										# 팝업창 글자수 세는 로직의 오동작 방지
			keyboard.send('enter')
			mouse.move(xPos, yPos)
			ctypes.windll.user32.BlockInput(False)				# 마우스/키보드 입력 전면 차단 해제

			time.sleep(0.5)


if __name__ == "__main__":
	#try:
		#os.chdir(sys._MEIPASS)
		#print(sys._MEIPASS)
	#except:
		#os.chdir(os.getcwd())
		#print(os.getcwd())
	app = Bypass()
