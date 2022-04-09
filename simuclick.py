import pyautogui
import time
import xlrd
import pyperclip
import datetime


config_path = 'config/dingdong/'
file = 'cmd.xls'
pyautogui.PAUSE = 0.03


def check_if_image_on_screen(img_path):
    global last_unix_time

    location = pyautogui.locateCenterOnScreen(img_path, confidence=0.8, grayscale=True, region=(0, 0, 620, 1104))
    if location is not None:
        print(f"{img_path}存在, now {datetime.datetime.now()}, delta time {time.time() - last_unix_time}")
        last_unix_time = time.time()
        return True
    else:
        print(f"{img_path}不存在, now {datetime.datetime.now()}, delta time {time.time() - last_unix_time}")
        last_unix_time = time.time()
        return False


def mouse_click(clickTimes, LR, img_path, reTry):
    if reTry == 1:
        while True:
            location = pyautogui.locateCenterOnScreen(img_path, confidence=0.8, grayscale=True, region=(0, 0, 620, 1104))
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, button=LR)
                break
            print("未找到匹配图片, 0.01秒后重试")
            time.sleep(0.01)
    elif reTry == -1:
        while True:
            location = pyautogui.locateCenterOnScreen(img_path, confidence=0.8, grayscale=True, region=(0, 0, 620, 1104))
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=LR)
            time.sleep(0.1)
    elif reTry > 1:
        i = 1
        while i < reTry + 1:
            location = pyautogui.locateCenterOnScreen(img_path, confidence=0.8, grayscale=True, region=(0, 0, 620, 1104))
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=LR)
                print("重复")
                i += 1
            time.sleep(0.1)


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
    cmd_cmd = True
    #行数检查
    if sheet1.nrows<2:
        print("没数据啊哥")
        cmd_cmd = False
    #每行数据检查
    i = 1
    while i < sheet1.nrows:
        # 第1列 操作类型检查
        cmd_type = sheet1.row(i)[0]
        if cmd_type.ctype != 2 or (cmd_type.value != 1.0 and cmd_type.value != 2.0 and cmd_type.value != 3.0
        and cmd_type.value != 4.0 and cmd_type.value != 5.0 and cmd_type.value != 6.0 and cmd_type.value != 7.0):
            print('第', i+1, "行,第1列数据有毛病")
            cmd_cmd = False
        # 第2列 内容检查
        cmd_value = sheet1.row(i)[1]
        # 读图点击类型指令，内容必须为字符串类型
        if cmd_type.value == 1.0 or cmd_type.value == 2.0 or cmd_type.value == 3.0:
            if cmd_value.ctype != 1:
                print('第', i+1, "行,第2列数据有毛病")
                cmd_cmd = False
        # 第3列 重复次数检查
        cmdTimes = sheet1.row(i)[2]
        # 内容必须为数字类型
        if cmdTimes.ctype not in [0, 2]:
            print('第', i+1, "行,第3列数据有毛病")
            cmd_cmd = False
        # 第4列 下一跳检查
        cmd_next = sheet1.row(i)[3]
        # 判断则为字符串，其余为数字
        if cmd_type.value == 7.0:
            if cmd_next.ctype != 1:
                print('第', i+1, "行,第4列数据有毛病")
                cmd_cmd = False
            else:
                if "," not in cmd_next.value:
                    print('第', i+1, "行,第4列数据有毛病")
                    cmd_cmd = False
        else:
            if cmd_next.ctype == 2 and ((cmd_next.value < 1 and cmd_next.value != 9999) or cmd_next.value >= sheet1.nrows):
                print('第', i+1, "行,第4列数据有毛病")
                cmd_cmd = False
            elif cmd_next.ctype != 2:
                print('第', i+1, "行,第4列数据有毛病")
                cmd_cmd = False

        # 输入类型，内容不能为空
        if cmd_type.value == 4.0:
            if cmd_value.ctype == 0:
                print('第', i+1, "行,第2列数据有毛病")
                cmd_cmd = False
        # 等待类型，内容必须为数字
        if cmd_type.value == 5.0:
            if cmd_value.ctype != 2:
                print('第', i+1, "行,第2列数据有毛病")
                cmd_cmd = False
        # 滚轮事件，内容必须为数字
        if cmd_type.value == 6.0:
            if cmd_value.ctype != 2:
                print('第', i+1, "行,第2列数据有毛病")
                cmd_cmd = False
        # 判断，内容不能为空，重复次数一定是空
        if cmd_type.value == 7.0:
            if cmd_value.ctype == 0:
                print('第', i+1, "行,第2列数据有毛病")
                cmd_cmd = False
            if cmdTimes.ctype != 0:
                print('第', i+1, "行,第2列数据有毛病")
                cmd_cmd = False
        i += 1
    return cmd_cmd


def task_handler(sheet):
    global last_unix_time
    next_row = 1
    while next_row != 9999:
        #取本行指令的操作类型
        print(f"拿下一行{next_row}, now {datetime.datetime.now()}, delta time {time.time() - last_unix_time}")
        last_unix_time = time.time()
        cmd_type = sheet.row(next_row)[0]
        _next_row = sheet.row(next_row)[3].value
        print(f"读完下一行{next_row}, now {datetime.datetime.now()}, delta time {time.time() - last_unix_time}")
        last_unix_time = time.time()
        if cmd_type.value == 1.0:
            #取图片名称
            img_name = sheet.row(next_row)[1].value
            reTry = 1
            if sheet.row(next_row)[2].ctype == 2 and sheet.row(next_row)[2].value != 0:
                reTry = sheet.row(next_row)[2].value
            mouse_click(1, "left", config_path + img_name, reTry)
            next_row = int(_next_row)
            print(f"单击左键{img_name}, 下一跳第{next_row}行")
        #2代表双击左键
        elif cmd_type.value == 2.0:
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
        elif cmd_type.value == 3.0:
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
            waitTime = sheet.row(next_row)[1].value
            time.sleep(waitTime)
            next_row = int(_next_row)
            print(f"等待{waitTime}秒, 下一跳第{next_row}行")
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
            print(f"开始判断{img_name}是否存在, now {datetime.datetime.now()}, delta time {time.time() - last_unix_time}")
            last_unix_time = time.time()
            if check_if_image_on_screen(img_path=config_path + img_name):
                str_next_row = _next_row.split(",")[0]
                print(f"判断{img_name}存在, now {datetime.datetime.now()}, delta time {time.time() - last_unix_time}")
            else:
                str_next_row = _next_row.split(",")[1]
                print(f"判断{img_name}不存在, now {datetime.datetime.now()}, delta time {time.time() - last_unix_time}")
            if str_next_row == "":
                next_row = 9999
            else:
                next_row = int(str_next_row)
            print(f"下一跳第{next_row}行")
        last_unix_time = time.time()


if __name__ == '__main__':
    # Code of your program here
    #打开文件
    last_unix_time = time.time()
    wb = xlrd.open_workbook(filename=config_path+file)
    #通过索引获取表格sheet页
    s = wb.sheet_by_index(0)
    #数据检查
    cmd_cmd = data_check(s)
    if cmd_cmd:
        task_handler(s)
    else:
        print('输入有误或者已经退出!')


