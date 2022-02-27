# import pyautogui
# import time
# time.sleep(10)
# img_path = '\"D:\\tmp\mylink.png\"'
# pyautogui.typewrite(img_path)
# time.sleep(10)
# pyautogui.press("enter")


import win32clipboard as w
import win32con
import win32api
import win32gui
import time
from apscheduler.schedulers.blocking import BlockingScheduler
import datetime
import os

def move_to_top(name = 'SSTap Beta 1.0.9.7 - 享受游戏'):
    def get_all_hwnd(hwnd, mouse):
        if (win32gui.IsWindow(hwnd) and
                win32gui.IsWindowEnabled(hwnd) and
                win32gui.IsWindowVisible(hwnd)):
            hwnd_map.update({hwnd: win32gui.GetWindowText(hwnd)})

    hwnd_map = {}
    win32gui.EnumWindows(get_all_hwnd, 0)

    for h, t in hwnd_map.items():
        print(h, t)
        if t:
            if t == name:
                # h 为想要放到最前面的窗口句柄
                print(h)

                win32gui.BringWindowToTop(h)
                # shell = win32com.client.Dispatch("WScript.Shell")
                # shell.SendKeys('%')

                # 被其他窗口遮挡，调用后放到最前面
                win32gui.SetForegroundWindow(h)

                # 解决被最小化的情况
                win32gui.ShowWindow(h, win32con.SW_RESTORE)


# 把文字放入剪贴板
def setText(aString):
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_UNICODETEXT, aString)
    w.CloseClipboard()


# 模拟ctrl+V
def ctrlV():
    win32api.keybd_event(17, 0, 0, 0)  # ctrl
    win32api.keybd_event(86, 0, 0, 0)  # V
    win32api.keybd_event(86, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键
    win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)


# 模拟alt+s
def altS():
    win32api.keybd_event(18, 0, 0, 0)
    win32api.keybd_event(83, 0, 0, 0)
    win32api.keybd_event(83, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(18, 0, win32con.KEYEVENTF_KEYUP, 0)


# 模拟enter
def enter():
    win32api.keybd_event(13, 0, 0, 0)
    win32api.keybd_event(13, 0, win32con.KEYEVENTF_KEYUP, 0)


# 模拟单击
def click():
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)


# 移动鼠标的位置
def movePos(x, y):
    win32api.SetCursorPos((x, y))
def log_string(log, string):
    message = str(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))+"  "+string
    log.write(message+ '\n')
    log.flush()
    print(message)

def create_path(dir):
    if not os.path.exists(dir):
        os.makedirs(dir)
def book_cmd():
    while True:
        try:
            book_cmd_old()
            break
        except:
            print("error")
            pass
    # book_cmd_old()
def book_cmd_old():
    today = str(time.strftime("%Y-%m-%d", time.localtime()))
    log_path = 'logs'
    create_path(log_path)
    log = open(log_path + "/{}_log.txt".format(today),"a")
    exe_path = r"C:\Program Files (x86)\Tencent\WeChat\WeChat.exe"
    win32api.ShellExecute(0, 'open', exe_path, '', '', 1)
    move_to_top("微信")
    # exit()

    # 获取鼠标当前位置
    # hwnd=win32gui.FindWindow("MozillaWindowClass",None)
    hwnd = win32gui.FindWindow("WeChatMainWndForPC", None)
    win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
    win32gui.MoveWindow(hwnd, 0, 0, 1000, 700, True)
    time.sleep(0.01)
    # 1.移动鼠标到通讯录位置，单击打开通讯录
    movePos(28, 147)
    click()
    # 2.移动鼠标到搜索框，单击，输入要搜索的名字
    movePos(148, 35)
    click()
    # setText('李顺友')
    # user_name = "郑富荣"
    # user_name = "小号"

    today = datetime.datetime.today()
    day_of_week = today.weekday()
    if day_of_week in [0]:
        user_name = "李顺友"
    else:
        user_name = "文件传输助手"
    log_string(log, user_name)
    setText(user_name)
    ctrlV()
    time.sleep(1)  # 别问我为什么要停1秒，问就是给微信一个反应的时间，他反应慢反应不过来，其他位置暂停的原因同样
    enter()
    time.sleep(1)
    # 3.复制要发送的消息，发送
    text_msg = "友哥,我想订周三20点到22点的一号场地,姓名:郑富荣,手机号:15217975717"
    setText(text_msg)
    log_string(log, text_msg)
    ctrlV()
    while True:
        now_time = datetime.datetime.now()
        cur_time = now_time.hour * 3600 + now_time.minute * 60 + now_time.second + now_time.microsecond / 1000000  # microsecond是微妙,转换为秒
        log_string(log, str(now_time))
        # 等待时间
        # 6500-M2,这个时间估计每台电脑都不一样?
        # if cur_time >= 14 * 3600 + 0.9:
        if cur_time >= 8 * 3600 + 0*60 + 0.5:
            # time.sleep(0.2)#休眠200MS，防止时间对不上
            # driver.find_element_by_class_name("btn-primary").click()  # 提交审核
            now_time = datetime.datetime.now()
            log_string(log, "click" + str(now_time))

            break
        else:
            time.sleep(0.1)
            log_string(log, "waiting...")
    altS()
    return 1

if __name__ == "__main__":
    code_version = 202111221400
    print(os.getcwd())
    print("file name: %s" % (__file__), ", code Version: ", code_version)

    scheduler = BlockingScheduler()
    # scheduler.add_job(func=book_cmd, args=(),
    #               trigger='cron', day_of_week=1 - 1, hour=7, minute=59, misfire_grace_time=60*5)  # 周1约周3
    scheduler.add_job(func=book_cmd, args=(),
                      trigger='cron', hour=7, minute=59, misfire_grace_time=60 * 5)  # 周1约周3
    scheduler.start()
    # book_cmd()
