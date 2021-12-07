import pyautogui
import time
import xlrd
import pyperclip


# 定义鼠标事件
# duration类似于移动时间或移动速度，省略后则是瞬间移动到指定的位置
def Mouse(click_times, img_name, retry_times):
    if retry_times == 1:
        location = pyautogui.locateCenterOnScreen(img_name, confidence=0.9)
        if location is not None:
            pyautogui.click(location.x, location.y, clicks=click_times, duration=0.2, interval=0.2)

    elif retry_times == -1:
        while True:
            location = pyautogui.locateCenterOnScreen(img_name,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=click_times, duration=0.2, interval=0.2)
    elif retry_times > 1:
        i = 1
        while i < retry_times + 1:
            location = pyautogui.locateCenterOnScreen(img_name,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=click_times, duration=0.2, interval=0.2)
                print("重复{}第{}次".format(img_name, i))
                i = i + 1

# cell_value     1.0：左键单击
#                2.0：输入字符串
#                3.0：等待
#                4.0：热键

# cell_type      空：0
#                字符串：1
#                数字：2
#                日期：3
#                布尔：4
#                error：5

# 任务一：进行一轮抢课
def WorkFunction1(sheet):
    i = 1
    while i < sheet.nrows:
        # 取excel表格中第i行操作
        cmd_type = sheet.cell_value(i, 1)
        # 1：左键单击
        if cmd_type == 1.0:
            # 获取图片名称
            img_name = sheet.cell_value(i, 2)
            retry_times = 1
            if sheet.cell_type(i, 3) == 2 and sheet.cell_value(i, 3) != 0:
                retry_times = sheet.cell_value(i, 3)
            Mouse(1, img_name, retry_times)
            print("单击左键:{}  Done".format(img_name))

        # 2：输入字符串
        elif cmd_type == 2.0:
            string = sheet.cell_value(i, 2)
            pyperclip.copy(string)
            pyautogui.hotkey('ctrl','v')
            print("输入字符串:{}  Done".format(string))
        # 3：等待
        elif cmd_type == 3.0:
            wait_time = sheet.cell_value(i, 2)
            time.sleep(wait_time)
            print("等待 {} 秒  Done".format(wait_time))
        # 4：键盘热键
        elif cmd_type == 4.0:
            hotkey = sheet.cell_value(i, 2)
            # 防止刷新过快停留在原网页
            time.sleep(1)
            pyautogui.hotkey(hotkey)
            print("按下 {}  Done".format(hotkey))
            time.sleep(1)
        i = i + 1

# 任务二：蹲点等人退课
def WorkFunction2(sheet) :
    while True:
        WorkFunction1(sheet)
        time.sleep(2)


if __name__ == '__main__':
    start_time = time.time()
    file = "info.xlsx"
    # 打开文件
    xr = xlrd.open_workbook(filename=file)
    # 通过索引顺序获取表单
    sheet = xr.sheet_by_index(0)
    print("------欢迎使用自动抢课脚本------")
    print("---------@danteking---------")
    print("1.抢课一次")
    print("2.蹲点等人退课后抢指定课")
    choice = input(">>")
    start_time = time.time()

    if choice == "1":
        WorkFunction1(sheet)
    elif choice == "2":
        WorkFunction2(sheet)
    else:
        print("非法输入，退出")
    end_time = time.time()
    time_consume = end_time - start_time
    time_consume = ('%.2f' % time_consume)
    print("耗时 {} 秒".format(time_consume))