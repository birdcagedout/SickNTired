from time import time
from PIL import ImageGrab, Image
from mss import mss
import d3dshot
start = time()

# 1번
#img = ImageGrab.grab()

# 2번
#with mss() as sct:
#	monitor = sct.monitors[2]
#	sct_img = sct.grab(monitor)
#	img = Image.frombytes('RGB', sct_img.size, sct_img.bgra, 'raw', 'BGRX')
#	img.save("1.png")



# 3번
#d = d3dshot.create()
#d.screenshot()


print(time() - start)