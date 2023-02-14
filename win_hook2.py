# 쓸데없는 프로그램이 가장 위(foreground)로 뜰 때
# 제목에 target 단어가 있으면 강제 종료
# https://blog.naver.com/PostView.naver?blogId=minim83&logNo=221963958883&parentCategoryNo=&categoryNo=13&viewDate=&isShowPopularPosts=true&from=search

import sys, time, ctypes, ctypes.wintypes, logging
import win32gui, win32con, win32api, win32process
#import pywinauto as pw

# Set logger
logging.basicConfig(format="[%(asctime)s]{%(filename)s:%(lineno)d}-%(levelname)s - %(message)s")
logger = logging.getLogger("My Logger")
logger.setLevel(logging.INFO)

def run():
	user32 = ctypes.windll.user32
	ole32 = ctypes.windll.ole32
	ole32.CoInitialize(0)

	WinEventProcType = ctypes.WINFUNCTYPE(None, ctypes.wintypes.HANDLE, ctypes.wintypes.DWORD, ctypes.wintypes.HWND, ctypes.wintypes.LONG, ctypes.wintypes.LONG, ctypes.wintypes.DWORD, ctypes.wintypes.DWORD)

	# Define callback function
	def callback(hWinEventHook, event, hwnd, idObject, idChild, dwEventThread, dwmsEventTIme):
		title = win32gui.GetWindowText(hwnd)

		#if title != "":
		#	logger.info(f"[{title}]")

		# Force to exit application
		#target_words = ["저공해", "계산기"]
		#for word in target_words:
		#	if word in title:
		#		win32gui.SendMessage(hwnd, win32con.WM_CLOSE, 0, 0)

		if "자동차관리정보시스템" in title:
			win_rect = win32gui.GetWindowRect(hwnd)
			print(f"[{int(time.time())}] Foreground: <<{title}>>")


	# Set eventhook
	def set_eventhook(WinEventProc, eventType):
		return user32.SetWinEventHook(
			eventType,
			eventType,
			0,
			WinEventProc,
			0,
			0,
			win32con.WINEVENT_OUTOFCONTEXT
		)
	WinEventProc = WinEventProcType(callback)
	user32.SetWinEventHook.restype = ctypes.wintypes.HANDLE

	# List events to catch
	events = [win32con.EVENT_SYSTEM_FOREGROUND, win32con.EVENT_OBJECT_FOCUS]
	hooks = [set_eventhook(WinEventProc, event) for event in events]
	#if not any(hooks):
	#	logger.info('SetWinEventHook failed')
	#	sys.exit(1)

	msg = ctypes.wintypes.MSG()
	while user32.GetMessageW(ctypes.byref(msg), 0, 0, 0) > 0:
		try:
			user32.TranslateMessageW(msg)
		except Exception as e:
			print(e)
		user32.DispatchMessageW(msg)

	# 메시지 루프 탈출시: WM_CLOSE
	for hook in hooks:
		user32.UnhookWinEvent(hook)

	ole32.CoUninitialize()

if __name__ == "__main__":
	run()