import os
import gc								# Garbage Collection
import sys
import time
import traceback
import pyperclip
import pyautogui
import keyboard
import ctypes
from tkinter import *
from tkinter import messagebox
from threading import *
import win32gui, win32ui, win32con


pyautogui.FAILSAFE = False		# pyautogui.FAILSAFE = True가 기본값

WIN_WIDTH = 350
WIN_HEIGHT = 350

WAITING_WIN_WIDTH = 200
WAITING_WIN_HEIGHT = 200

BG_COLOR = 	'#dae5f7'	# '#D7ccf5'= 연한보라

ACTIVATED = False
THREAD_EXPIRED = False
THREAD_DEAD = False

SCREEN_WIDTH = 1920
SCREEN_HEIGHT = 1080

REGION = (300, 300, 1200, 400)

HOME_PATH = os.path.dirname(__file__)
VER = "2.99u"							# Pil/Grab OSError + keyboard 모듈


# 인트로 스플래시 이미지 닫기
import importlib
if '_PYIBoot_SPLASH' in os.environ and importlib.util.find_spec("pyi_splash"):
	import pyi_splash
	#pyi_splash.update_text('UI Loaded ...')
	pyi_splash.close()


# prevent the computer from going to sleep
ctypes.windll.kernel32.SetThreadExecutionState(0x80000003)


# 디스플레이 DPI 조정 보정
#user32 = ctypes.windll.user32
#user32.SetProcessDPIAware()


# 간혹 프로그램 원인 모르게 죽는 이유 ==> PIL에서 grab 호출시 OSError 발생
# ==> 줄줄이 raise OSError 
#Fail to allocate bitmap
#Traceback (most recent call last):
#  File "Bypass299b.py", line 393, in run
#  File "pyautogui\__init__.py", line 175, in wrapper
#  File "pyautogui\__init__.py", line 207, in locateCenterOnScreen
#  File "pyscreeze\__init__.py", line 413, in locateCenterOnScreen
#  File "pyscreeze\__init__.py", line 372, in locateOnScreen
#  File "pyscreeze\__init__.py", line 145, in wrapper
#  File "pyscreeze\__init__.py", line 457, in _screenshot_win32
#  File "PIL\ImageGrab.py", line 46, in grab
#OSError: screen grab failed


class Bypass:
	def __init__(self):
		# tkinter
		self.root = Tk()
		self.root.iconbitmap(os.path.join(HOME_PATH, "tube.ico"))
		self.root.geometry(f"{WIN_WIDTH}x{WIN_HEIGHT}+{int((self.root.winfo_screenwidth()-WIN_WIDTH)/2-3)}+{int((self.root.winfo_screenheight()-WIN_HEIGHT)/2-3)}")
		self.root.title(f"귀찮다 v{VER}")
		self.root.resizable(False, False)
		self.root.configure(bg=BG_COLOR)
		#self.root.wm_attributes("-topmost", True)

		# "법인번호" 텍스트 라벨
		self.label_0 = Label(self.root, text="자동 입력할 조회 사유를 선택하세요", font=("Malgun Gothic", 15, 'bold'))
		self.label_0.configure(bg=BG_COLOR)

		self.radio_var = IntVar()
		self.reasons = ["", "신규등록 등 업무 대상자 확인", "이전등록 등 업무 대상자 확인", "보험 가입 여부 확인", "소유자 확인"]	
		self.radio_btn1 = Radiobutton(self.root, text="신규등록 등 업무 대상자 확인", font=("Malgun Gothic", 12), value=1, variable=self.radio_var, command=self.on_select)
		self.radio_btn2 = Radiobutton(self.root, text="이전등록 등 업무 대상자 확인", font=("Malgun Gothic", 12), value=2, variable=self.radio_var, command=self.on_select)
		self.radio_btn3 = Radiobutton(self.root, text="보험 가입 여부 확인", font=("Malgun Gothic", 12), value=3, variable=self.radio_var, command=self.on_select)
		self.radio_btn4 = Radiobutton(self.root, text="소유자 확인", font=("Malgun Gothic", 12), value=4, variable=self.radio_var, command=self.on_select)
		#self.radio_btn5 = Radiobutton(self.root, text="", font=("Malgun Gothic", 12), value=5, variable=self.radio_var, command=self.on_select)

		self.checkButtonVar = IntVar(value=1)
		self.checkButton = Checkbutton(self.root, variable = self.checkButtonVar, bg=BG_COLOR)
		self.checkButton["state"] = DISABLED

		self.label_1 = Label(self.root, text="투명도 ", font=("Malgun Gothic", 11, 'bold'), fg="#FDAA00")
		self.scale = Scale(self.root, from_=0, to=90, resolution=10, showvalue=0, bd=2, orient=HORIZONTAL, length=180,
							activebackground='#358CD6', highlightbackground=BG_COLOR, troughcolor="#FDDB71", command=self.updateScale)
		#self.scale.bind("<ButtonRelease-1>", self.updateScale)
		
		#self.entry_5 = Entry(self.root, font=("Malgun Gothic", 12), fg="#707070", width=25)
		#self.entry_5.insert(0, "직접 입력(한글 잘 안 되지롱~)")
		#self.entry_5.bind("<Button-1>", self.on_input)

		self.radio_btn1.configure(bg=BG_COLOR)
		self.radio_btn2.configure(bg=BG_COLOR)
		self.radio_btn3.configure(bg=BG_COLOR)
		self.radio_btn4.configure(bg=BG_COLOR)
		#self.radio_btn5.configure(bg=BG_COLOR)
		self.label_1.configure(bg=BG_COLOR)
		self.scale.configure(bg=BG_COLOR)

		self.label_0.grid(row=0, column=0, sticky=EW, padx=0, pady=10, columnspan=1)
		self.radio_btn1.grid(row=1, column=0, sticky=W, padx=35, pady=0)
		self.radio_btn2.grid(row=2, column=0, sticky=W, padx=35, pady=0)
		self.radio_btn3.grid(row=3, column=0, sticky=W, padx=35, pady=0)
		self.radio_btn4.grid(row=4, column=0, sticky=W, padx=35, pady=0)
		#self.radio_btn5.grid(row=5, column=0, sticky=W, padx=35, pady=0)
		#self.entry_5.grid(row=5, column=0, sticky=W, padx=62, pady=0)
		self.checkButton.grid(row=5, column=0, sticky=W, padx=(35, 0), pady=0)
		self.label_1.grid(row=5, column=0, sticky=W, padx=(57,0), pady=0)
		self.scale.grid(row=5, column=0, sticky=W, padx=(110, 0), columnspan=2, pady=3)

		# 버튼
		self.btn_img = PhotoImage(file=os.path.join(HOME_PATH, "btn1.png"))
		self.start_btn = Button(self.root, image=self.btn_img, borderwidth=1)
		self.start_btn.bind("<ButtonRelease-1>", self.on_start)
		self.start_btn.configure(width=300, height=99, activebackground="#FFFFFF")
		self.start_btn.grid(row=6, column=0, sticky=EW, padx=25, pady=18, columnspan=1)

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
		self.target_win_hwnd = None
		self.target_win_title = None
		
		self.get_target_win()
		if self.target_win_hwnd == None:
			messagebox.showerror("오류 발생", "자동차관리정보시스템이 시작되지 않았습니다.\n프로그램을 종료합니다.")
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

		# Identifiers
		gwl_style = -16

		ws_minimizebox = 131072
		ws_maximizebox = 65536

		swp_nozorder = 4
		swp_nomove = 2
		swp_nosize = 1
		swp_framechanged = 32

		hwnd = get_parent(root.winfo_id())

		old_style = get_window_long(hwnd, gwl_style) # Get the style
		new_style = old_style & ~ ws_maximizebox & ~ ws_minimizebox # New style, without max/min buttons

		set_window_long(hwnd, gwl_style, new_style) # Apply the new style
		set_window_pos(hwnd, 0, 0, 0, 0, 0, swp_nomove | swp_nosize | swp_nozorder | swp_framechanged)     # Updates


	# 현재 (보이는 윈도 중에서) 원하는 윈도 찾기 : "자동차관리정보시스템" 문자열 포함
	def get_target_win(self):
		def winEnumHandler(hwnd, ctx):
			if win32gui.IsWindowVisible(hwnd):
				title = win32gui.GetWindowText(hwnd)
				if "자동차관리정보시스템" in title:
					self.target_win_hwnd = hwnd
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
			messagebox.showerror("오류 발생", "자동 입력할 사유를 선택해주세요.")
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

		distance_x = scr_width - (pos_x + WIN_WIDTH) + 143		# X거리(여유분=50)
		distance_y = scr_height - (pos_y + WIN_HEIGHT) + 81  	# Y거리(여유분=100)
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
		self.root.update()

		# 윈도 크기를 오리크기(WAITING_WIN_WIDTH)로 줄임
		while self.root.winfo_width() > WAITING_WIN_WIDTH:
			self.root.geometry(f"{self.root.winfo_width() - 5}x{self.root.winfo_height() - 5}")
			self.root.update()
		
		# 작은 윈도 모드 + 투명
		self.root.title(f"귀찮다 v{VER} 작동 중")
		#self.root.wm_attributes('-toolwindow', True)
		#self.root.wm_attributes("-alpha", 0.2)

		# 오리 스마일 세팅
		i = 0
		image_on_canvas = self.canvas.create_image(0, 0, anchor=NW, image=self.smile_frames[i])
		#self.canvas.pack(fill="both", expand=True)
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
				#self.root.wm_attributes("-topmost", True)
				self.root.update()
				ctypes.windll.kernel32.SetThreadExecutionState(0x80000003)

			# exception 발생 ==> 쓰레드 종료
			global THREAD_DEAD
			if THREAD_DEAD == True:
				THREAD_DEAD = False
				if self.watcher.is_alive() == True:
					self.watcher.join()
				del self.watcher
				self.watcher = WatcherThread(reason = self.reason)
				self.watcher.start()

			# 팝업창 이미지 발견!!
			global ACTIVATED
			if ACTIVATED == True:
				ACTIVATED = False

				# 쓰레드 자연사 시키고 ==> 오리 빡침 돌리고 ==> 빡침 끝나면 새로운 쓰레드
				global THREAD_EXPIRED
				THREAD_EXPIRED = True

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
					time.sleep(0.02)

				if self.watcher.is_alive() == True:
					self.watcher.join()
				del self.watcher
				#print("join 후 쓰레드 종료됨")
				self.watcher = WatcherThread(reason = self.reason)
				self.watcher.start()
			
			time.sleep(0.1)


class WatcherThread(Thread):
	# 초기화
	def __init__(self, reason=[]):
		super(WatcherThread, self).__init__(daemon=True)
		self.reason = reason

	def run(self):
		# 2.98에서 추가 ==> 전원관리 timer 초기화
		timer = 0
		#start_t = time.time()

		while True:
			global THREAD_DEAD

			# 이미지 위치변수 초기화
			pos1 = None
			gab1 = None

			# 팝업창 이미지 찾기
			try:
				# 연계정보 조회
				pos1 = pyautogui.locateCenterOnScreen(os.path.join(HOME_PATH,"inquiry_popup.PNG"), grayscale=True, region=REGION)
				#pos2 = locateCenterOnScreen("inquiry_save.PNG", grayscale=True)

				# 갑부 열람사유입력
				if pos1 == None:
					gab1 = pyautogui.locateCenterOnScreen(os.path.join(HOME_PATH, "gabbu.PNG"), grayscale=True, region=REGION)
					

			except OSError as ose:
				#print(f"[locateCenterOnScreen 에러] 종류:{type(ose)}\t 내용:{ose}")	# 종류=Exception, 내용 = traceback.print_exc()과 동일
				#traceback.print_exc()													# Traceback (most recent call last)
				#global pyautogui
				#importlib.reload(pyautogui)
				THREAD_DEAD = True
				break

			except Exception as e:
				#print(f"[locateCenterOnScreen 에러] 종류:{type(e)}\t 내용:{e}")		# 종류=Exception, 내용 = traceback.print_exc()과 동일
				#traceback.print_exc()													# Traceback (most recent call last)
				#time.sleep(2)
				#importlib.reload(pyautogui)
				THREAD_DEAD = True
				break

			# 두개 모두 발견 못한 경우 ==> loop 재시작
			if (pos1 == None) and (gab1 == None):
				time.sleep(0.3)
				continue
			# 둘 중 하나는 발견한 경우
			else:
				# 작동중 ==> 메인 쓰레드에 알림 
				global ACTIVATED
				ACTIVATED = True

				# 연계정보 조회인 경우
				if pos1 != None:
					# 조회 사유 자동입력
					pos2 = pyautogui.Point(x=pos1.x+410, y=pos1.y+40)
					#pos2 = Point(x=pos1.x+510, y=pos1.y+50)	# 강선임 주임님
					pos1 = pyautogui.Point(x=pos1.x, y=pos1.y+80)
				
				# 갑부 열람사유입력인 경우
				if gab1 != None:
					pos2 = pyautogui.Point(x=gab1.x+130, y=gab1.y+35)
					pos1 = pyautogui.Point(x=gab1.x, y=gab1.y+100)
				
			# 사유 입력창 클릭
			pyautogui.moveTo(pos1)
			pyautogui.click()

			# 사유 입력
			try:
				keyboard.write(self.reason)
			except:
				pass

			# 저장 클릭
			pyautogui.moveTo(pos2)
			pyautogui.click()
			#now = dt.fromtimestamp(time.time()).strftime('%Y-%m-%d %H:%M:%S')
			#print(f"[{now}] 저장 버튼 클릭 완료!")

			#===============================================================
			# 2.98에서 추가 ==> sleep 방지
			if timer < sys.maxsize:
				timer += 1
			else:
				timer = 0
			if timer % 300 == 0:		# 대충 157 = 60초
				#print(f"{timer}까지 실행 시간: {time.time() - start_t}")
				ctypes.windll.kernel32.SetThreadExecutionState(0x80000003)
				try:
					gc.collect()
				except:
					pass

			# 2.98에서 추가(끝)
			#===============================================================
			
			# 1초 쉬기
			time.sleep(1)

			# 한번 실행되었다면 ==> 메인쓰레드에서 현재쓰레드 종료시킴
			global THREAD_EXPIRED
			if THREAD_EXPIRED == True:
				THREAD_EXPIRED = False
				break


if __name__ == "__main__":
	#try:
		#os.chdir(sys._MEIPASS)
		#print(sys._MEIPASS)
	#except:
		#os.chdir(os.getcwd())
		#print(os.getcwd())
	app = Bypass()
