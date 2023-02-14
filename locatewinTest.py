import win32gui
from time import *
#import pyautogui
import keyboard
import mouse
import ctypes

import win32con, win32api


while True:
	start = time()
	try:
		hWnd = win32gui.GetForegroundWindow()
		title = win32gui.GetWindowText(hWnd)
		rect = win32gui.GetWindowRect(hWnd)
	except Exception:
		continue
	x = rect[0]
	y = rect[1]
	w = rect[2] - x
	h = rect[3] - y

	#print(f"X: {x}, Y: {y}, width: {w}, height: {h}")

	if title != "조회 사유 입력" and title != "개인정보 공개 열람사유입력":
		sleep(0.1)
		continue
	#elif title == "조회 사유 입력":
	#	pos2 = pyautogui.Point(x + 480, y + 55)
	#elif title == "개인정보 공개 열람사유입력":
	#	pos2 = pyautogui.Point(x + 300, y + 55)
	
	else:
		childhWnd = win32gui.GetWindow(hWnd, win32con.GW_CHILD)
		while childhWnd != 0:
			edithWnd = win32gui.GetWindow(hWnd, )
		win32api.SendMessage(edithWnd, win32con.WM_CHAR, ord('H'), 0)
		win32api.Sleep(500)

		#pos1 = pyautogui.Point(x + 110, y + int(h*0.45))
		#ctypes.windll.user32.BlockInput(True)				# 마우스/키보드 입력 전면 차단
		#pyautogui.moveTo(x + 110, y + int(h*0.45))

		#pyautogui.moveTo(x + 110, y + 110)
		#pyautogui.click()
		#mouse.move(x + 110, y + 110)
		#mouse.click()

		#keyboard.write("신규등록 등 업무 대상자 확인")
		#keyboard.send('enter')
		#ctypes.windll.user32.BlockInput(False)				# 마우스/키보드 입력 전면 차단 해제
		#print(f"elapsed time: {time() - start}")
		#sleep(0.1)

