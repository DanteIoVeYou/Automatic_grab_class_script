import pyautogui
import time
import xlrd
import pyperclip

# 定义鼠标事件

def Mouse(click_times, x_pos, y_pos):
    pyautogui.click(x=x_pos, y=y_pos, clicks=click_times, interval= 0.15, duration= 0.1, button="left")  # move to 100, 200, then click the left mouse button.

def WorkFunction(sheet):
    i = 1
    # 读取操作类型
    while i < sheet.nrows:
        cmd_type = sheet.cell_value(i, 3)
        if cmd_type == 1.0:
            # 取每一个带点击目标的坐标
            x_pos = sheet.cell_value(i, 1)
            y_pos = sheet.cell_value(i, 2)
            print("第{}步：x={},y={} Done".format(i, x_pos, y_pos))
            Mouse(1, x_pos, y_pos)
        elif cmd_type == 0.0:
            # 三秒防刷/等待加载
            sleep_time = sheet.cell_value(i, 4)
            time.sleep(sleep_time)
            sleep_time = ('%.2f' % sleep_time)
            print("等待 {} 秒 Done".format(sleep_time))
        elif cmd_type == 2.0:
            input_value = sheet.cell_value(i, 4)
            pyperclip.copy(input_value)
            pyautogui.hotkey('ctrl', 'v')
            print("输入: {} Done".format(input_value))
        i = i + 1
if __name__ == '__main__':
    file = "info.xlsx"
    #打开文件
    xr = xlrd.open_workbook(filename=file)
    #通过索引顺序获取
    sheet = xr.sheet_by_index(0)
    print("------欢迎使用自动抢课脚本------")
    print("---------@danteking---------")
    print("1.开始")
    choice = input(">>")
    start_time = time.time()
    if choice == "1":
        WorkFunction(sheet)
        # cmd_type = sheet.cell_value(2, 3)
        # print(cmd_type)

    else:
        print("非法输入，退出")
    end_time = time.time()
    time_consume = end_time - start_time
    time_consume = ('%.2f' % time_consume)
    print("累计用时 {} 秒".format(time_consume))



