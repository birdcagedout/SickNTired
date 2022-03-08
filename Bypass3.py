import os
import cv2
import sys
import time
import pyautogui
import pyperclip
import numpy as np
from tkinter import *
from pyautogui import *
from threading import *
import win32gui, win32ui, win32con

WIN_WIDTH = 350
WIN_HEIGHT = 350

WAITING_WIN_WIDTH = 200
WAITING_WIN_HEIGHT = 200

SCREEN_WIDTH = 1920
SCREEN_HEIGHT = 1080

BG_COLOR = 	'#dae5f7'	# '#D7ccf5'= 연한보라

ACTIVATED = False

HOME_PATH = os.path.dirname(__file__)
VER = "3.0"



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

		self.target_img = None
		self.target_img_width = 0
		self.target_img_height = 0
		self.method = None

		self.list_win_names()

		# 윈도 크기 재조정
		#self.refresh_win_position()

	# 스크린샷 찍어서 이미지 리턴
	def get_screenshot(self):
		#hwnd = None	# 전체
		#hwnd = win32gui.FindWindow(None, self.target_name)
		wDC = win32gui.GetWindowDC(self.target_hwnd)
		dcObj = win32ui.CreateDCFromHandle(wDC)
		cDC = dcObj.CreateCompatibleDC()
		dataBitMap = win32ui.CreateBitmap()
		dataBitMap.CreateCompatibleBitmap(dcObj, self.screen_width, self.screen_height)
		cDC.SelectObject(dataBitMap)
		cDC.BitBlt((0,0), (self.screen_width, self.screen_height), dcObj, (0,0), win32con.SRCCOPY)
		#dataBitMap.SaveBitmapFile(cDC, bmpfilenamename)

		signedIntArray = dataBitMap.GetBitmapBits(True)
		img = np.fromstring(signedIntArray, dtype='uint8')
		img.shape = (self.screen_height, self.screen_width, 4)
		img = img[..., :3]
		img = np.ascontiguousarray(img)

		# Free Resources
		dcObj.DeleteDC()
		cDC.DeleteDC()
		win32gui.ReleaseDC(self.target_hwnd, wDC)
		win32gui.DeleteObject(dataBitMap.GetHandle())

		return img

	# 윈도 크기 재조정
	def refresh_win_position(self):
		win_rect = win32gui.GetWindowRect(self.target_hwnd)
		if win_rect[0] < 0:
			win32gui.MoveWindow(self.target_hwnd, 0, win_rect[1], (win_rect[2] - win_rect[0]), win_rect[3], True)
		if win_rect[2] > 1920:
			win32gui.MoveWindow(self.target_hwnd, (win_rect[0]), win_rect[1], (win_rect[2] - win_rect[0]), win_rect[3], True)
		
		win_rect = win32gui.GetWindowRect(self.target_hwnd)
		self.screen_width = win_rect[2] - win_rect[0]		#     0      1    2    3
		self.screen_height = win_rect[3] - win_rect[1]		# (X시작, Y시작, X끝, Y끝)
		self.offset_x = win_rect[0]
		self.offset_y = win_rect[1]
		#print(f"win_rect: {win_rect[0]}, {win_rect[1]}")


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
		for process_name in self.process_list:
			if win_name in process_name:
				self.target_name = process_name
				index = self.process_list.index(process_name)
				self.target_hwnd = self.hwnd_list[index]
				#print(f"win_name: {win_name}, self.target_name: {self.target_name}, self.target_hwnd: {self.target_hwnd}")
			else:
				#print("[get_target_win] 해당 이름의 process를 찾을 수 없습니다.")
				pass

	# 절대 위치 찾기
	def get_absolute_position(self, pos):
		return (pos[0] + self.offset_x, pos[1] + self.offset_y)

	
	# target 이미지 설정
	def set_target(self, target_img, method=cv2.TM_CCOEFF_NORMED):
		# load the image we're trying to match
		# https://docs.opencv.org/4.2.0/d4/da8/group__imgcodecs.html
		target_img_path = os.path.join(HOME_PATH, target_img)
		self.target_img = cv2.imread(target_img_path, cv2.IMREAD_GRAYSCALE)

		# Save the dimensions of the needle image
		self.target_img_width = self.target_img.shape[1]
		self.target_img_height = self.target_img.shape[0]

		# There are 6 methods to choose from:
		# TM_CCOEFF, TM_CCOEFF_NORMED, TM_CCORR, TM_CCORR_NORMED, TM_SQDIFF, TM_SQDIFF_NORMED
		self.method = method

	# target 이미지 찾기
	def find(self, haystack_img, threshold=0.5, debug_mode=None):
		# run the OpenCV algorithm
		result = cv2.matchTemplate(haystack_img, self.target_img, self.method)

		# Get the all the positions from the match result that exceed our threshold
		locations = np.where(result >= threshold)
		locations = list(zip(*locations[::-1]))
		#print(locations)

		# You'll notice a lot of overlapping rectangles get drawn. We can eliminate those redundant
		# locations by using groupRectangles().
		# First we need to create the list of [x, y, w, h] rectangles
		rectangles = []
		for loc in locations:
			rect = [int(loc[0]), int(loc[1]), self.target_img_width, self.target_img_height]
			# Add every box to the list twice in order to retain single (non-overlapping) boxes
			rectangles.append(rect)
			rectangles.append(rect)
		# Apply group rectangles.
		# The groupThreshold parameter should usually be 1. If you put it at 0 then no grouping is
		# done. If you put it at 2 then an object needs at least 3 overlapping rectangles to appear
		# in the result. I've set eps to 0.5, which is:
		# "Relative difference between sides of the rectangles to merge them into a group."
		rectangles, weights = cv2.groupRectangles(rectangles, groupThreshold=1, eps=0.5)
		#print(rectangles)

		points = []
		if len(rectangles):
			#print('Found needle.')

			line_color = (0, 255, 0)
			line_type = cv2.LINE_4
			marker_color = (255, 0, 255)
			marker_type = cv2.MARKER_CROSS

			# Loop over all the rectangles
			for (x, y, w, h) in rectangles:

				# Determine the center position
				center_x = x + int(w/2)
				center_y = y + int(h/2)
				# Save the points
				points.append(Point(center_x, center_y))

				if debug_mode == 'rectangles':
					# Determine the box position
					top_left = (x, y)
					bottom_right = (x + w, y + h)
					# Draw the box
					cv2.rectangle(haystack_img, top_left, bottom_right, color=line_color, lineType=line_type, thickness=2)
				elif debug_mode == 'points':
					# Draw the center point
					cv2.drawMarker(haystack_img, (center_x, center_y), color=marker_color, markerType=marker_type, markerSize=40, thickness=2)

		if debug_mode:
			cv2.imshow('Matches', haystack_img)
			#cv.waitKey()
			#cv.imwrite('result_click_point.jpg', haystack_img)

		return points


app = WindowCapture()
app.get_target_win("자동차관리정보시스템")
app.set_target("inquiry_popup.PNG")
#print(app.target_hwnd, app.target_name)
#while True:
#	app.refresh_win_position()
#exit()

while True:
	# 성능 측정
	time_base = time.time()

	# pyautogui로 스크린샷 찍기
	screenshot = app.get_screenshot()
	#screenshot = np.array(screenshot)

	app.find(screenshot, threshold=0.6, debug_mode="rectangles")

	#cv2.imshow("Computer Vison", screenshot)

	print(f"FPS: {1 / (time.time() - time_base)}")		# 사무실: 전체화면 10~12 FPS, 창만 30FPS

	# q를 누르면 exit
	#if cv2.waitKey(1) == ord('q'):
	#	cv2.destroyAllWindows()
	#	break

print('Done.')