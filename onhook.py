import sys
import time
import traceback
import win32con
import win32gui
from ctypes import windll, byref, c_ubyte, c_long
from ctypes.wintypes import HWND
from datetime import datetime

import cv2
import numpy as np

PostMessageW = windll.user32.PostMessageW
GetDC = windll.user32.GetDC
CreateCompatibleDC = windll.gdi32.CreateCompatibleDC
GetClientRect = windll.user32.GetClientRect
CreateCompatibleBitmap = windll.gdi32.CreateCompatibleBitmap
SelectObject = windll.gdi32.SelectObject
BitBlt = windll.gdi32.BitBlt
GetBitmapBits = windll.gdi32.GetBitmapBits
DeleteObject = windll.gdi32.DeleteObject
ReleaseDC = windll.user32.ReleaseDC

SRCCOPY = 0x00CC0020
WM_LBUTTONDOWN = 0x0201
WM_LBUTTONUP = 0x202
MAX_SIMI = 0.99


def left_down(handle: HWND, x: int, y: int):
    # https://docs.microsoft.com/en-us/windows/win32/inputdev/wm-lbuttondown
    wparam = 0
    lparam = y << 16 | x
    PostMessageW(handle, WM_LBUTTONDOWN, wparam, lparam)


def left_up(handle: HWND, x: int, y: int):
    # https://docs.microsoft.com/en-us/windows/win32/inputdev/wm-lbuttonup
    wparam = 0
    lparam = y << 16 | x
    PostMessageW(handle, WM_LBUTTONUP, wparam, lparam)


def set_current_window(hwnd):
    if win32gui.IsIconic(hwnd):
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    else:
        win32gui.SetForegroundWindow(hwnd)


def capture(handle: HWND):
    set_current_window(handle)
    time.sleep(1.0)
    left, top, right, bot = win32gui.GetWindowRect(handle)
    width = right - left
    height = bot - top
    # 开始截图
    dc = GetDC(0)
    cdc = CreateCompatibleDC(dc)
    bitmap = CreateCompatibleBitmap(dc, width, height)
    SelectObject(cdc, bitmap)
    BitBlt(cdc, 0, 0, width, width, dc, left, top, SRCCOPY)
    # 截图是BGRA排列，因此总元素个数需要乘以4
    total_bytes = width * height * 4
    # 使用 bytearray 来创建一个缓冲区
    buffer = bytearray(total_bytes)
    # 创建一个 c_ubyte 数组的类型
    byte_array = c_ubyte * total_bytes
    # 获取指向 byte_array 的指针
    pointer_to_byte_array = byref(byte_array.from_buffer(buffer))
    GetBitmapBits(bitmap, c_long(total_bytes), pointer_to_byte_array)
    DeleteObject(bitmap)
    DeleteObject(cdc)
    ReleaseDC(handle, dc)
    # 返回截图数据为numpy.ndarray
    return np.frombuffer(buffer, dtype=np.uint8).reshape(height, width, 4)


def match(img: [], path: str):
    gray_img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    temp = cv2.imread(path, 0)
    # 模板匹配，将alpha作为mask，TM_CCORR_NORMED方法的计算结果范围为[0, 1]，越接近1越匹配
    result = cv2.matchTemplate(gray_img, temp, cv2.TM_CCORR_NORMED)
    return cv2.minMaxLoc(result), temp


def img_match(handle: HWND):
    path = 'img/'
    filenames = ['talent1.png', 'talent2.png', 'talent3.png']
    img = capture(handle)
    for filename in filenames:
        (val_min, val_max, min_loc, max_loc), temp = match(img, path + filename)
        # print("相似度", val_max)
        if val_max < MAX_SIMI:
            return False
    return True


def button_pos(handle: HWND):
    img = capture(handle)
    (val_min, val_max, min_loc, max_loc), temp = match(img, "img/button.png")
    if val_max < MAX_SIMI:
        raise Exception("button not found")
    width, height = temp.shape[::-1]
    top_left = max_loc
    bottom_right = top_left[0] + width, top_left[1] + height
    return top_left[0] + (width >> 1), top_left[1] + (height >> 1)


def except_hook(tp, val, tb):
    f_error = open("except_error.log", 'a')
    trace_list = traceback.format_tb(tb)
    html = repr(tp) + "\n"
    html += (repr(val) + "\n")
    for line in trace_list:
        html += (line + "\n")
    # print(html, file=sys.stderr)
    print(datetime.now(), file=f_error)
    print(html, file=f_error)
    f_error.close()


def main():
    WINDOW_NAME = u'射雕  '
    try:
        # 防止UI放大导致截图不完整
        windll.user32.SetProcessDPIAware()
        handle = windll.user32.FindWindowW(None, WINDOW_NAME)
        x, y = button_pos(handle)
        print(x, y)

        while not img_match(handle):
            left_down(handle, x - 10, y - 45)
            time.sleep(0.1)
            left_up(handle, x - 10, y - 45)
            time.sleep(3.5)
    except Exception:
        raise


if __name__ == "__main__":
    sys.excepthook = except_hook
    if not windll.shell32.IsUserAnAdmin():
        # 不是管理员就提权
        windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
    else:
        main()
    # main()
