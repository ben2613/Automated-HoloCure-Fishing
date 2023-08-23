import collections
from PIL import Image
import win32con, win32gui, win32ui
import cv2
import numpy as np

# stuff for bg screenshot
hwnd = 0
wRect = None
Box = collections.namedtuple('Box', 'left top width height')

def init_screenshot_bg_var(hwnd_p):
    global hwnd, wRect
    hwnd = hwnd_p
    wRect = win32gui.GetWindowRect(hwnd)

def screenshot(region):
    im = capture(region=region)
    return im

def matchTemplate(cv_im_0, cv_im_1, confidence):
    w, h = cv_im_1.shape[:2]
    res = cv2.matchTemplate(cv_im_0, cv_im_1, cv2.TM_CCOEFF_NORMED)
    loc = np.where(res >= confidence)
    for pt in zip(*loc[::-1]):  # Switch columns and rows
        # return the first match
        return Box(pt[0], pt[1], w, h)
    return None

def locate(path, pic, grayscale, confidence):
    cv_haystack_im = cv2.cvtColor(np.array(pic), cv2.COLOR_BGR2GRAY if grayscale else cv2.COLOR_BGR2RGB)
    cv_needle_im = cv2.imread(path, cv2.IMREAD_GRAYSCALE if grayscale else None)
    return matchTemplate(cv_haystack_im, cv_needle_im, confidence)

def locateOnScreen(path, grayscale, region=None, confidence=0.6):
    im = capture(region)
    return locate(path, im, grayscale, confidence)

def capture(region=None):
    global hwnd, wRect
    if not hwnd or not wRect:
        return False
    
    w, h = 0, 0
    srcPos = (0, 0)
    if region is None:
        w, h = wRect[2] - wRect[0], wRect[3] - wRect[1]
    else:
        w, h = region[2], region[3]
        srcPos = (region[0] - wRect[0], region[1] - wRect[1])
    wDC = win32gui.GetWindowDC(hwnd)
    dcObj = win32ui.CreateDCFromHandle(wDC)
    cDC = dcObj.CreateCompatibleDC()
    dataBitMap = win32ui.CreateBitmap()
    dataBitMap.CreateCompatibleBitmap(dcObj, w, h)
    cDC.SelectObject(dataBitMap)
    cDC.BitBlt((0, 0), (w, h), dcObj, srcPos, win32con.SRCCOPY)
    # bmpinfo = dataBitMap.GetInfo()
    bmpstr = dataBitMap.GetBitmapBits(True)
    # Free Resources
    dcObj.DeleteDC()
    cDC.DeleteDC()
    win32gui.ReleaseDC(hwnd, wDC)
    win32gui.DeleteObject(dataBitMap.GetHandle())
    return Image.frombuffer('RGB', (w, h), bmpstr, 'raw', 'BGRX', 0, 1)
