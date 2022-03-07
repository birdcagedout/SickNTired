import cv2
import numpy as np
import os
import time
import pyautogui
import win32gui, win32ui, win32con



# 원하는 윈도를 캡쳐하는 클래스
class WindowCapture:
	# properties
	screen_width = 1920
	screen_height = 1080

	def __init__(self):
		self.screen_width = 1920	# screen_width
		self.screen_height = 1080	# screen_height
		self.offset_x = 0
		self.offset_y = 0

		self.hwnd_list = []
		self.process_list = []

		self.target_hwnd = None
		self.target_name = None

		self.list_win_names()

	def get_screenshot(self):

		#hwnd = None	# 전체
		hwnd = win32gui.FindWindow(None, self.target_name)
		wDC = win32gui.GetWindowDC(hwnd)
		dcObj = win32ui.CreateDCFromHandle(wDC)
		cDC = dcObj.CreateCompatibleDC()
		dataBitMap = win32ui.CreateBitmap()
		dataBitMap.CreateCompatibleBitmap(dcObj, self.screen_width, self.screen_height)
		cDC.SelectObject(dataBitMap)
		cDC.BitBlt((0,0),(self.screen_width, self.screen_height) , dcObj, (0,0), win32con.SRCCOPY)
		#dataBitMap.SaveBitmapFile(cDC, bmpfilenamename)

		signedIntArray = dataBitMap.GetBitmapBits(True)
		img = np.fromstring(signedIntArray, dtype='uint8')
		img.shape = (self.screen_height, self.screen_width, 4)
		img = img[..., :3]
		img = np.ascontiguousarray(img)

		# Free Resources
		dcObj.DeleteDC()
		cDC.DeleteDC()
		win32gui.ReleaseDC(hwnd, wDC)
		win32gui.DeleteObject(dataBitMap.GetHandle())

		# 윈도 크기 재조정
		self.refresh_win_position()

		return img

	# 윈도 크기 재조정
	def refresh_win_position(self):
		win_rect = win32gui.GetWindowRect(self.target_hwnd)
		if win_rect[0] < 0:
			win32gui.MoveWindow(self.target_hwnd, 0, win_rect[1], (win_rect[2] - win_rect[0]), win_rect[3], True)
		
		win_rect = win32gui.GetWindowRect(self.target_hwnd)
		self.screen_width = win_rect[2] - win_rect[0]		#     0      1    2    3
		self.screen_height = win_rect[3] - win_rect[1]		# (X시작, Y시작, X끝, Y끝)
		self.offset_x = win_rect[0]
		self.offset_y = win_rect[1]
		print(f"win_rect: {win_rect[0]}, {win_rect[1]}")


	# 현재 (보이는 윈도 중에서) 원하는 윈도 찾기 : "자동차관리정보시스템" 문자열 포함
	def list_win_names(self):
		def winEnumHandler(hwnd, ctx):
			if win32gui.IsWindowVisible(hwnd):
				print(hex(hwnd), win32gui.GetWindowText(hwnd))
				self.hwnd_list.append(hwnd)
				self.process_list.append(win32gui.GetWindowText(hwnd))
		win32gui.EnumWindows(winEnumHandler, None)


	# 원하는 윈도만 찾기
	def get_target_win(self, win_name):
		for index, process_name in enumerate(self.process_list):
			if win_name in process_name:
				self.target_name = process_name
				self.target_hwnd = self.hwnd_list[index]
				#print(f"win_name: {win_name}, self.target_name: {self.target_name}, self.target_hwnd: {self.target_hwnd}")
			else:
				#print("[get_target_win] 해당 이름의 process를 찾을 수 없습니다.")
				pass

	# 절대 위치 찾기
	def get_absolute_position(self, pos):
		return (pos[0] + self.offset_x, pos[1] + self.offset_y)


app = WindowCapture()
app.get_target_win("자동차관리정보시스템")
#print(app.target_hwnd, app.target_name)
while True:
	app.refresh_win_position()
exit()

while True:
	# 성능 측정
	time_base = time.time()

	# pyautogui로 스크린샷 찍기
	screenshot = app.get_screenshot()
	screenshot = np.array(screenshot)
	#screenshot = cv2.cvtColor(screenshot, cv2.COLOR_RGB2BGR)

	cv2.imshow("Computer Vison", screenshot)

	print(f"FPS: {1 / (time.time() - time_base)}")		# 사무실: 전체화면 10~12 FPS, 창만 30FPS

	# q를 누르면 exit
	if cv2.waitKey(1) == ord('q'):
		cv2.destroyAllWindows()
		break

print('Done.')