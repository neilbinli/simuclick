import pyautogui
import time
import keyboard
import xlrd
import pyperclip
import datetime
import ctypes
import ctypes.wintypes
from python_imagesearch.imagesearch import *
from PIL import Image

config_path = 'config/dingdong/'
file = 'cmd.xls'
pyautogui.PAUSE = 0.03
working_region = (0, 0, 620, 1104)
im = region_grabber(working_region)
confidence_level = 0.8


def check_if_image_on_screen(img_path, is_refresh_test_region=True):
    global last_unix_time, working_region, confidence_level
    if is_refresh_test_region:
        location = imagesearcharea(img_path, *working_region, precision=confidence_level)
    else:
        location = imagesearcharea(img_path, *working_region, precision=confidence_level, im=im)
    image = Image.open(img_path)
    image_w, image_h = image.size
    if location[0] != -1:
        location = ctypes.wintypes.POINT(int(location[0]+image_w/2), int(location[1]+image_h/2))
        print(f"{img_path}存在({location.x}, {location.y}), now {datetime.datetime.now()}, delta time {time.time() - last_unix_time}")
        last_unix_time = time.time()
        return location
    else:
        print(f"{img_path}不存在, now {datetime.datetime.now()}, delta time {time.time() - last_unix_time}")
        last_unix_time = time.time()
        return None


def mouse_click(click_times, LR, img_path, is_refresh_test_region=True):
    while True:
        location = check_if_image_on_screen(img_path, is_refresh_test_region=is_refresh_test_region)
        if location is not None:
            pyautogui.click(location.x, location.y, clicks=click_times, button=LR)
            break
        print("未找到匹配图片, 0.01秒后重试")
        time.sleep(0.01)


def data_check(sheet1):
    '''
    # 数据检查
    # cmd_type.value  1.0 左键单击    2.0 左键双击  3.0 右键单击  4.0 输入  5.0 等待  6.0 滚轮
    # ctype     空：0
    #           字符串：1
    #           数字：2
    #           日期：3
    #           布尔：4
    #           error：5
    '''
    check_result = True
    #行数检查
    if sheet1.nrows<2:
        print("没数据啊哥")
        check_result = False
    #每行数据检查
    i = 1
    while i < sheet1.nrows:
        # 第1列 操作类型检查
        cmd_type = sheet1.row(i)[0]
        if cmd_type.ctype != 2 or (cmd_type.value != 1.0 and cmd_type.value != 2.0 and cmd_type.value != 3.0
        and cmd_type.value != 4.0 and cmd_type.value != 5.0 and cmd_type.value != 6.0 and cmd_type.value != 7.0):
            print('第', i+1, "行,第1列数据有毛病")
            check_result = False
        # 第2列 内容检查
        cmd_value = sheet1.row(i)[1]
        # 读图点击类型指令，内容必须为字符串类型
        if cmd_type.value == 1.0 or cmd_type.value == 2.0 or cmd_type.value == 3.0:
            if cmd_value.ctype != 1:
                print('第', i+1, "行,第2列数据有毛病")
                check_result = False
        # 第3列 是否获取新界面（默认1获取，0不获取）
        is_refresh_test_region = sheet1.row(i)[2]
        # 内容必须为数字类型
        if is_refresh_test_region.ctype not in [0, 2]:
            print('第', i+1, "行,第3列数据有毛病")
            check_result = False
        if is_refresh_test_region.ctype == 2 and is_refresh_test_region.value not in [0, 1.0]:
            print('第', i+1, "行,第3列数据有毛病")
            check_result = False
        # 第4列 下一跳检查
        cmd_next = sheet1.row(i)[3]
        # 判断则为字符串，其余为数字
        if cmd_type.value == 7.0:
            if cmd_next.ctype != 1:
                print('第', i+1, "行,第4列数据有毛病")
                check_result = False
            else:
                if "," not in cmd_next.value:
                    print('第', i+1, "行,第4列数据有毛病")
                    check_result = False
        else:
            if cmd_next.ctype == 2 and ((cmd_next.value < 1 and cmd_next.value != 9999) or cmd_next.value >= sheet1.nrows):
                print('第', i+1, "行,第4列数据有毛病")
                check_result = False
            elif cmd_next.ctype != 2:
                print('第', i+1, "行,第4列数据有毛病")
                check_result = False

        # 输入类型，内容不能为空
        if cmd_type.value == 4.0:
            if cmd_value.ctype == 0:
                print('第', i+1, "行,第2列数据有毛病")
                check_result = False
        # 等待类型，内容必须为数字
        if cmd_type.value == 5.0:
            if cmd_value.ctype != 2:
                print('第', i+1, "行,第2列数据有毛病")
                check_result = False
        # 滚轮事件，内容必须为数字
        if cmd_type.value == 6.0:
            if cmd_value.ctype != 2:
                print('第', i+1, "行,第2列数据有毛病")
                check_result = False
        # 判断，内容不能为空，重复次数一定是空
        if cmd_type.value == 7.0:
            if cmd_value.ctype == 0:
                print('第', i+1, "行,第2列数据有毛病")
                check_result = False
        i += 1
    return check_result


def task_handler(sheet):
    global last_unix_time
    next_row = 1
    while next_row != 9999 and keyboard.is_pressed('q') == False:
        #取本行指令的操作类型
        last_unix_time = time.time()
        cmd_type = sheet.row(next_row)[0]
        _next_row = sheet.row(next_row)[3].value
        last_unix_time = time.time()
        #取是否获取新画面
        is_refresh_test_region = True
        if sheet.row(next_row)[2].ctype == 2 and sheet.row(next_row)[2].value == 0:
            is_refresh_test_region = False
        if cmd_type.value == 1.0:
            #取图片名称
            img_name = sheet.row(next_row)[1].value
            mouse_click(1, "left", config_path + img_name, is_refresh_test_region=is_refresh_test_region)
            next_row = int(_next_row)
            print(f"单击左键{img_name}, 下一跳第{next_row}行")
        #2代表双击左键
        elif cmd_type.value == 2.0:
            #取图片名称
            img_name = sheet.row(next_row)[1].value
            mouse_click(2, "left", config_path + img_name, is_refresh_test_region=is_refresh_test_region)
            next_row = int(_next_row)
            print(f"双击左键{img_name}, 下一跳第{next_row}行")
        #3代表右键
        elif cmd_type.value == 3.0:
            #取图片名称
            img_name = sheet.row(next_row)[1].value
            mouse_click(1, "right", config_path + img_name, is_refresh_test_region=is_refresh_test_region)
            next_row = int(_next_row)
            print(f"右键{img_name}, 下一跳第{next_row}行")
        #4代表输入
        elif cmd_type.value == 4.0:
            #取输入值
            inputValue = sheet.row(next_row)[1].value
            pyperclip.copy(inputValue)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)
            next_row = int(_next_row)
            print(f"输入:{inputValue}, 下一跳第{next_row}行")
        #5代表等待
        elif cmd_type.value == 5.0:
            #取等待时间
            wait_time = sheet.row(next_row)[1].value
            time.sleep(wait_time)
            next_row = int(_next_row)
            print(f"等待{wait_time}秒, 下一跳第{next_row}行")
        #6代表滚轮
        elif cmd_type.value == 6.0:
            #取滚动长度
            scroll = sheet.row(next_row)[1].value
            pyautogui.scroll(int(scroll))
            next_row = int(_next_row)
            print(f"滚轮滑动{int(scroll)}距离, 下一跳第{next_row}行")
        #7代表判断
        elif cmd_type.value == 7.0:
            #取图片名称
            img_name = sheet.row(next_row)[1].value
            last_unix_time = time.time()
            if check_if_image_on_screen(img_path=config_path + img_name, is_refresh_test_region=is_refresh_test_region):
                str_next_row = _next_row.split(",")[0]
            else:
                str_next_row = _next_row.split(",")[1]
            if str_next_row == "":
                next_row = 9999
            else:
                next_row = int(str_next_row)
            print(f"下一跳第{next_row}行")
        last_unix_time = time.time()


if __name__ == '__main__':
    #打开文件
    last_unix_time = time.time()
    wb = xlrd.open_workbook(filename=config_path+file)
    #通过索引获取表格sheet页
    s = wb.sheet_by_index(0)
    #数据检查
    check_result = data_check(s)
    if check_result:
        task_handler(s)
    else:
        print('输入有误或者已经退出!')


