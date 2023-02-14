import ctypes
import time
from ctypes import wintypes
wintypes.ULONG_PTR = wintypes.WPARAM
hllDll = ctypes.WinDLL ("User32.dll", use_last_error=True)
VK_HANGUEL = 0x15

class MOUSEINPUT(ctypes.Structure):
    _fields_ = (("dx",          wintypes.LONG),
                ("dy",          wintypes.LONG),
                ("mouseData",   wintypes.DWORD),
                ("dwFlags",     wintypes.DWORD),
                ("time",        wintypes.DWORD),
                ("dwExtraInfo", wintypes.ULONG_PTR))

class HARDWAREINPUT(ctypes.Structure):
    _fields_ = (("uMsg",    wintypes.DWORD),
                ("wParamL", wintypes.WORD),
                ("wParamH", wintypes.WORD))

class KEYBDINPUT(ctypes.Structure):
    _fields_ = (("wVk",         wintypes.WORD),
                ("wScan",       wintypes.WORD),
                ("dwFlags",     wintypes.DWORD),
                ("time",        wintypes.DWORD),
                ("dwExtraInfo", wintypes.ULONG_PTR))

class INPUT(ctypes.Structure):
    class _INPUT(ctypes.Union):
        _fields_ = (("ki", KEYBDINPUT),
                    ("mi", MOUSEINPUT),
                    ("hi", HARDWAREINPUT))
    _anonymous_ = ("_input",)
    _fields_ = (("type",   wintypes.DWORD),
                ("_input", _INPUT))

def get_hanguel_state():
    return hllDll.GetKeyState(VK_HANGUEL) & 0x0001
	#return hllDll.GetAsyncKeyState(VK_HANGUEL)
	# GetKeyState의 리턴값은 0xffffff81, 0xffffff80, 0x00000001, 0x00000000
	# 65409=0xFF81, 

def change_state():
    x = INPUT(type=1 ,ki=KEYBDINPUT(wVk=VK_HANGUEL))
    y = INPUT(type=1, ki=KEYBDINPUT(wVk=VK_HANGUEL,dwFlags=2))
    hllDll.SendInput(1, ctypes.byref(x), ctypes.sizeof(x))
    time.sleep(0.05)
    hllDll.SendInput(1, ctypes.byref(y), ctypes.sizeof(y))

while True:
	time.sleep(2)
	#한 > 영
	if get_hanguel_state() == 1: #1 일경우 vk_key : 0x15(한글키)가 활성화
		#change_state() #한글키 누르고(key_press) , 때기(release)
		print("한글 상태입니다.")

	#영 > 한
	if get_hanguel_state() == 0: #0 일경우 vk_key : 0x15(한글키)가 비활성화
		#change_state() #한글키 누르고(key_press) , 때기(release)
		print("영문 상태입니다.")