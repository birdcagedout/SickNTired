import time
from pyautogui import *

while True:
	# "민원열람 > 갑부열람"
	pos1 = locateCenterOnScreen("gabbu.PNG", grayscale=True)
	time.sleep(0.5)

	if pos1 == None:
		time.sleep(0.5)
		continue
	else:
		# 같은 등록번호 차량이 여러대인 경우 해당차량 선택 ==> "자동차등록원부 조회"
		while True:
			pos2 = locateCenterOnScreen("wonbu_select.PNG", grayscale=True)
			if pos2 != None:
				time.sleep(0.5)
				continue
			else:
				break
			
		#press(['tab', 'tab', 'tab', 'tab', 'tab', 'end'])
		time.sleep(0.5)

		press(['tab', 'tab', 'tab', 'tab', 'tab', 'up'])
					
		press('f4')