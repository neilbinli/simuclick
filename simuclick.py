import pyautogui
import time
import xlrd
import pyperclip
import pydirectinput
from pynput.mouse import Listener, Controller, Button
import win32api, win32con
import mouse
import ctypes, sys


config_path = 'config/dingdong/'
file = 'cmd.xls'


def check_if_image_on_screen(img_path):
    location = pyautogui.locateCenterOnScreen(img_path, confidence=0.8)
    if location is not None:
        return False
    else:
        return True


def mouse_click(clickTimes, LR, img_path, reTry):
    if reTry == 1:
        while True:
            location = pyautogui.locateCenterOnScreen(img_path, confidence=0.8)
            if location is not None:
                pyautogui.moveTo(location.x, location.y)
                pyautogui.click(location.x, location.y, clicks=clickTimes, button=LR)
                break
            print("未找到匹配图片, 0.1秒后重试")
            time.sleep(0.1)
    elif reTry == -1:
        while True:
            location = pyautogui.locateCenterOnScreen(img_path, confidence=0.8)
            if location is not None:
                pyautogui.move(location.x, location.y)
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=LR)
            time.sleep(0.1)
    elif reTry > 1:
        i = 1
        while i < reTry + 1:
            location = pyautogui.locateCenterOnScreen(img_path, confidence=0.8)
            if location is not None:
                pyautogui.move(location.x, location.y)
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=LR)
                print("重复")
                i += 1
            time.sleep(0.1)


def data_check(sheet1):
    '''
    # 数据检查
    # cmdType.value  1.0 左键单击    2.0 左键双击  3.0 右键单击  4.0 输入  5.0 等待  6.0 滚轮
    # ctype     空：0
    #           字符串：1
    #           数字：2
    #           日期：3
    #           布尔：4
    #           error：5
    '''
    checkCmd = True
    #行数检查
    if sheet1.nrows<2:
        print("没数据啊哥")
        checkCmd = False
    #每行数据检查
    i = 1
    while i < sheet1.nrows:
        # 第1列 操作类型检查
        cmdType = sheet1.row(i)[0]
        if cmdType.ctype != 2 or (cmdType.value != 1.0 and cmdType.value != 2.0 and cmdType.value != 3.0
        and cmdType.value != 4.0 and cmdType.value != 5.0 and cmdType.value != 6.0 and cmdType.value != 7.0):
            print('第', i+1, "行,第1列数据有毛病")
            checkCmd = False
        # 第2列 内容检查
        cmdValue = sheet1.row(i)[1]
        # 读图点击类型指令，内容必须为字符串类型
        if cmdType.value == 1.0 or cmdType.value == 2.0 or cmdType.value == 3.0:
            if cmdValue.ctype != 1:
                print('第', i+1, "行,第2列数据有毛病")
                checkCmd = False
        # 第3列 重复次数检查
        cmdTimes = sheet1.row(i)[2]
        # 内容必须为数字类型
        if cmdTimes.ctype not in [0, 2]:
            print('第', i+1, "行,第3列数据有毛病")
            checkCmd = False
        # 第4列 下一跳检查
        cmdNext = sheet1.row(i)[3]
        # 判断则为字符串，其余为数字
        if cmdType.value == 7.0:
            if cmdNext.ctype != 1:
                print('第', i+1, "行,第4列数据有毛病")
                checkCmd = False
            else:
                if "," not in cmdNext.value:
                    print('第', i+1, "行,第4列数据有毛病")
                    checkCmd = False
        else:
            if cmdNext.ctype == 2 and ((cmdNext.value < 1 and cmdNext.value != -1) or cmdNext.value >= sheet1.nrows):
                print('第', i+1, "行,第4列数据有毛病")
                checkCmd = False
            elif cmdNext.ctype != 2:
                print('第', i+1, "行,第4列数据有毛病")
                checkCmd = False

        # 输入类型，内容不能为空
        if cmdType.value == 4.0:
            if cmdValue.ctype == 0:
                print('第', i+1, "行,第2列数据有毛病")
                checkCmd = False
        # 等待类型，内容必须为数字
        if cmdType.value == 5.0:
            if cmdValue.ctype != 2:
                print('第', i+1, "行,第2列数据有毛病")
                checkCmd = False
        # 滚轮事件，内容必须为数字
        if cmdType.value == 6.0:
            if cmdValue.ctype != 2:
                print('第', i+1, "行,第2列数据有毛病")
                checkCmd = False
        # 判断，内容不能为空，重复次数一定是空
        if cmdType.value == 7.0:
            if cmdValue.ctype == 0:
                print('第', i+1, "行,第2列数据有毛病")
                checkCmd = False
            if cmdTimes.ctype != 0:
                print('第', i+1, "行,第2列数据有毛病")
                checkCmd = False
        i += 1
    return checkCmd


def task_handler(sheet):
    next_row = 1
    while next_row != -1:
        #取本行指令的操作类型
        cmdType = sheet.row(next_row)[0]
        _next_row = sheet.row(next_row)[3].value
        if cmdType.value == 1.0:
            #取图片名称
            img_name = sheet.row(next_row)[1].value
            reTry = 1
            if sheet.row(next_row)[2].ctype == 2 and sheet.row(next_row)[2].value != 0:
                reTry = sheet.row(next_row)[2].value
            mouse_click(1, "left", config_path + img_name, reTry)
            next_row = int(_next_row)
            print(f"单击左键{img_name}, 下一跳第{next_row}行")
        #2代表双击左键
        elif cmdType.value == 2.0:
            #取图片名称
            img_name = config_path + sheet.row(next_row)[1].value
            #取重试次数
            reTry = 1
            if sheet.row(next_row)[2].ctype == 2 and sheet.row(next_row)[2].value != 0:
                reTry = sheet.row(next_row)[2].value
            mouse_click(2, "left", config_path + img_name, reTry)
            next_row = int(_next_row)
            print(f"双击左键{img_name}, 下一跳第{next_row}行")
        #3代表右键
        elif cmdType.value == 3.0:
            #取图片名称
            img_name = sheet.row(next_row)[1].value
            #取重试次数
            reTry = 1
            if sheet.row(next_row)[2].ctype == 2 and sheet.row(next_row)[2].value != 0:
                reTry = sheet.row(next_row)[2].value
            mouse_click(1, "right", config_path + img_name, reTry)
            next_row = int(_next_row)
            print(f"右键{img_name}, 下一跳第{next_row}行")
        #4代表输入
        elif cmdType.value == 4.0:
            #取输入值
            inputValue = sheet.row(next_row)[1].value
            pyperclip.copy(inputValue)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)
            next_row = int(_next_row)
            print(f"输入:{inputValue}, 下一跳第{next_row}行")
        #5代表等待
        elif cmdType.value == 5.0:
            #取等待时间
            waitTime = sheet.row(next_row)[1].value
            time.sleep(waitTime)
            next_row = int(_next_row)
            print(f"等待{waitTime}秒, 下一跳第{next_row}行")
        #6代表滚轮
        elif cmdType.value == 6.0:
            #取滚动长度
            scroll = sheet.row(next_row)[1].value
            pyautogui.scroll(int(scroll))
            next_row = int(_next_row)
            print(f"滚轮滑动{int(scroll)}距离, 下一跳第{next_row}行")
        #7代表判断
        elif cmdType.value == 7.0:
            #取图片名称
            img_name = sheet.row(next_row)[1].value
            if check_if_image_on_screen(img_path=config_path + img_name):
                str_next_row = _next_row.split(",")[0]
            else:
                str_next_row = _next_row.split(",")[1]
            if str_next_row == "":
                next_row = -1
            else:
                next_row = int(str_next_row)


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


if __name__ == '__main__':
    if is_admin():
        # Code of your program here
        #打开文件
        wb = xlrd.open_workbook(filename=config_path+file)
        #通过索引获取表格sheet页
        s = wb.sheet_by_index(0)
        #数据检查
        checkCmd = data_check(s)
        if checkCmd:
            task_handler(s)
        else:
            print('输入有误或者已经退出!')
    else:
        # Re-run the program with admin rights
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)


