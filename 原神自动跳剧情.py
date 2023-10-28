import ctypes
import sys
import time
from threading import Thread

import cv2
import keyboard
import win32api
import win32con
import win32gui
from PIL import ImageGrab


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception as e:
        return False


class genshin_control():
    def __init__(self):
        self.loc = []
        self.size = []
        self.hwnd = win32gui.FindWindow("UnityWndClass", "原神")
        self.auto_template = cv2.imread('auto.png')
        self.talk_template = cv2.imread('talk.png')
        self.status = False
        self.get_loc()
        win32gui.SendMessage(self.hwnd, win32con.WM_SYSCOMMAND, win32con.SC_RESTORE, 0)
        win32gui.SetForegroundWindow(self.hwnd)

    def get_loc(self):

        rect = win32gui.GetWindowRect(self.hwnd)
        self.loc = []
        self.loc.append(rect[0])
        self.loc.append(rect[1])
        self.loc.append(rect[2])
        self.loc.append(rect[3])
        self.size = []
        self.size.append(rect[2] - rect[0])
        self.size.append(rect[3] - rect[1])

    def mouse_move(self, x, y):
        ctypes.windll.user32.SetCursorPos(x, y)

    def mouse_up(self, x, y):
        time.sleep(0.1)
        xx, yy = x + self.loc[0], y + self.loc[1]
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, xx, yy, 0, 0)

    def mouse_down(self, x, y):
        time.sleep(0.1)
        xx, yy = x + self.loc[0], y + self.loc[1]
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, xx, yy, 0, 0)

    def mouse_click(self, x, y):
        self.mouse_move(x, y)
        self.mouse_down(x, y)
        self.mouse_up(x, y)

    def match_img(self, img, target, type=cv2.TM_CCOEFF_NORMED):
        '''

        :param img: 截屏文件
        :param target: template模板
        :param type: None
        :return:
        '''
        h, w = target.shape[:2]
        res = cv2.matchTemplate(img, target, type)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
        if type in [cv2.TM_SQDIFF, cv2.TM_SQDIFF_NORMED]:
            return (
                *min_loc,
                min_loc[0] + w,
                min_loc[1] + h,
                min_loc[0] + w // 2,
                min_loc[1] + h // 2,
                min_val
            )
        else:
            return (
                *max_loc,
                max_loc[0] + w,
                max_loc[1] + h,
                max_loc[0] + w // 2,
                max_loc[1] + h // 2,
                max_val
            )

    def find_pic(self, target, name):
        max_loc = None
        win32gui.SendMessage(self.hwnd, win32con.WM_SYSCOMMAND, win32con.SC_RESTORE, 0)
        win32gui.SetForegroundWindow(self.hwnd)
        self.get_loc()
        img = ImageGrab.grab((self.loc[0], self.loc[1], self.loc[2], self.loc[3]))
        img.save('img.jpg')
        img = cv2.imread('img.jpg')
        if name == 'auto':
            img = img[0:200, 0:300]
        if name == 'talk':
            img = img[int(self.size[1] / 2):int(self.size[1]), int(self.size[0] / 2):int(self.size[0])]
        max_loc = self.match_img(img, target)
        '''(62, 73, 79, 95, 70, 84, 0.9934518337249756)
        0,1 找图左上位置 2，3找图右下位置 4，5找图中心位置 6 匹配度
            '''
        print(max_loc)

        return max_loc

    def istalking(self):
        max_loc = self.find_pic(self.auto_template, 'auto')
        if max_loc[6] > 0.82:
            # os.system("pause")
            print("匹配")
            return True
        else:
            max_loc = self.find_pic(self.talk_template, 'talk')
            if max_loc[6] > 0.85:
                print("匹配")
                return True
            else:
                print("不匹配")
                time.sleep(1)
                return False

    def skip_talk(self):
        while True:
            if self.status == False:
                break

            if self.istalking() == True:
                # self.mouse_click(100,600)
                # self.pyautogui_click(1079,642)
                x = int(self.loc[0] + 1092 / 1622 * self.size[0])
                y = int(self.loc[1] + 725 / 956 * self.size[1])
                print(x, y, self.loc, self.size)
                self.mouse_click(x, y)

    def thread(self):
        thread = Thread(target=self.skip_talk)
        thread.start()


if is_admin():
    gc = genshin_control()
    print('请最小化本窗口，按V键启动自动跳过')
    while True:
        keyboard.wait('v')
        gc.status = True
        print('自动跳过已启动，按V键关闭自动跳过')
        gc.thread()

        keyboard.wait('v')
        print('自动跳过已停止')
        gc.status = False

else:
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
