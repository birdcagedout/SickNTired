import os
import sys
import time
import pyperclip
from tkinter import *
from datetime import datetime as dt
from pyautogui import *

if __name__ == "__main__":

	# tkinter
	root = Tk()
	root.iconbitmap(r'corp_icon.ico')
	#root.geometry(f"{WIN_WIDTH}x{WIN_HEIGHT}+{int((self.root.winfo_screenwidth()-self.root.winfo_width())/2)}+{int((self.root.winfo_screenheight()-self.root.winfo_height())/2)}")
	root.title("귀찮다 v2.0")
	root.resizable(False, False)
	root.wm_attributes("-topmost", True)

	# "법인번호" 텍스트 라벨
	label_0 = Label(root, text=" 자동으로 입력할 조회 사유를 선택하세요.")
	label_0.grid(row=0, column=0, sticky=N, padx=0, pady=3)
	label_1 = Label(root, text=" 신규등록 등 업무 대상자 확인")
	label_1.grid(row=0, column=0, sticky=N, padx=0, pady=3)

	# 법인번호 입력칸
	self.entry_Regnum = Entry(self.root)
	self.entry_Regnum.grid(row=1, column=0, sticky=N, padx=10, pady=0 )
	self.entry_Regnum.bind("<KeyRelease>", self.verify_input)
	self.entry_Regnum.bind("<Return>", self.search)

	# "법인정보 열람" 버튼(이미지)
	self.btn_img = PhotoImage(file=r"res/btn_search2.png")
	self.btn_search = Button(self.root, image=self.btn_img, borderwidth=1, command=self.search)
	self.btn_search.grid(row=0, column=1, rowspan=2, sticky=S, padx=2, pady=5)
	self.btn_search["state"] = DISABLED

	# 법인번호 입력칸에 포커스
	self.entry_Regnum.focus()


	try:
		# 사유는 커맨드라인 아규먼트로 입력 받음
		reason = ""
		while len(sys.argv) == 1:
			print("====================================================")
			print("      자동으로 입력할 조회 사유를 선택하세요.      ")
			print("----------------------------------------------------")
			print("(1) 신규등록 등 업무 대상자 확인")
			print("(2) 이전등록 등 업무 대상자 확인")
			print("(3) 말소압류 등 업무 대상자 확인")
			print("(4) 보험 가입 여부 조회")
			print("(5) 직접 입력")
			print("----------------------------------------------------")
			choice = input("조회 사유(1~5) 선택: ").strip()[0]
			if choice == "1":
				reason = "신규등록 등 업무 대상자 확인"
				break
			elif choice == "2":
				reason = "이전등록 등 업무 대상자 확인"
				break
			elif choice == "3":
				reason = "말소압류 등 업무 대상자 확인"
				break
			elif choice == "4":
				reason = "보험 가입 여부 조회"
				break
			elif choice == "5":
				reason = input("사유 직접 입력: ").strip()
				break
			else:
				print("잘못된 입력입니다. 처음으로 돌아갑니다.")
				continue

		if len(sys.argv) >= 2:
			for idx, arg in enumerate(sys.argv):
				if idx >= 1:
					reason += arg + " "
		
		print(f"자동 입력 사유: {reason}")
		print("====================================================")
		print("")
		print("■■■■■■■■■■■■■■■■■■■■■■■■■■■")
		print("■     (경) 귀찮다 v1.8 작동을 시작합니다~ (축)     ■")
		print("■        프로그램 종료를 원하시면 : Ctrl + C       ■")
		print("■■■■■■■■■■■■■■■■■■■■■■■■■■■")
		print("")

		try:
			os.chdir(sys._MEIPASS)
			#print(sys._MEIPASS)
		except:
			os.chdir(os.getcwd())
			#print(os.getcwd())

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
			pyperclip.copy(reason)
			hotkey('ctrl', 'v')
			# 저장해둔 클립보드 내용을 현재 클립보드에 원래대로 복사해놓음
			pyperclip.copy(currentClipboard)
			now = dt.fromtimestamp(time.time()).strftime('%Y-%m-%d %H:%M:%S')
			print(f"[{now}] 사유 자동 입력 완료!")
			
			# 저장 클릭
			moveTo(pos2)
			click()
			now = dt.fromtimestamp(time.time()).strftime('%Y-%m-%d %H:%M:%S')
			print(f"[{now}] 저장 버튼 클릭 완료!")
			
			# 2초 쉬기
			print("")
			time.sleep(2)
	
	except KeyboardInterrupt:
		print("(깜딱이야!) 사용자 입력에 의해 프로그램을 종료합니다.")
		print("♩~ ♪~ ♬~ ♩~ ♪~ ♬~      ㅃㅏ2~ㅇㅕㅁ by 김재형")

		sys.exit(0)