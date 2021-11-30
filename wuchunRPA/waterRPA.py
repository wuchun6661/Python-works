import random

import pyautogui
import time
import xlrd
import pyperclip

yang = ["【小臭杨专属】【早上回复0】该上班拉！亲一口好好干活！mua~",
        "【小臭杨专属】【中午回复1】中午干饭不积极，上班没动力！",
        "【小臭杨专属】【晚上回复2】下班啦~臭臭杨，回去的时候注意安全~",
        "【小臭杨专属】【夜间回复3】收拾好了吗，快来让老公抱抱！"]
yyd = ["【yyd专属】【早上回复0】起这么早？太卷了吧",
       "【yyd专属】【中午回复1】yyd中午吃什么？Burger King？",
       "【yyd专属】【晚上回复2】到晚上啦，yyd，去喝淮南牛肉汤吧~",
       "【yyd专属】【夜间回复3】逼话少说，上号！"]
guo = ["【郭师兄专属】【早上回复0】早上好啊郭师兄",
       "【郭师兄专属】【中午回复1】该吃午饭啦，郭师兄",
       "【郭师兄专属】【晚上回复2】到晚上啦，郭师兄，早点吃完早点回来继续学习啊！",
       "【郭师兄专属】【夜间回复3】我先溜啦郭师兄~"]
time_dic = dict(morning=0, noon=1, evening=2, night=3)


# 定义鼠标事件
def mouseClick(clickTimes, lOrR, img, reTry):
    if reTry == 1:
        while True:
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2,
                                button=lOrR)  # interval间隔、duration持续时间
                break
            print("未找到匹配图片,0.1秒后重试")
            time.sleep(0.1)
    elif reTry == -1:
        while True:
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
            time.sleep(0.1)
    elif reTry > 1:
        i = 1
        while i < reTry + 1:
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
                print("重复")
                i += 1
            time.sleep(0.1)


# 数据检查
# cmdType.value  1.0 左键单击    2.0 左键双击  3.0 右键单击  4.0 输入  5.0 等待  6.0 滚轮
# ctype     空：0
#           字符串：1
#           数字：2
#           日期：3
#           布尔：4
#           error：5
def dataCheck(sheet1):
    checkCmd = True
    # 行数检查
    if sheet1.nrows < 2:
        print("没数据啊哥")
        checkCmd = False
    # 每行数据检查
    i = 1
    while i < sheet1.nrows:
        # 第1列 操作类型检查
        cmdType = sheet1.row(i)[0]
        if cmdType.ctype != 2 or (cmdType.value != 1.0 and cmdType.value != 2.0 and cmdType.value != 3.0
                                  and cmdType.value != 4.0 and cmdType.value != 5.0 and cmdType.value != 6.0):
            print('第', i + 1, "行,第1列数据有毛病")
            checkCmd = False
        # 第2列 内容检查
        cmdValue = sheet1.row(i)[1]
        # 读图点击类型指令，内容必须为字符串类型
        if cmdType.value == 1.0 or cmdType.value == 2.0 or cmdType.value == 3.0:
            if cmdValue.ctype != 1:
                print('第', i + 1, "行,第2列数据有毛病")
                checkCmd = False
        # 输入类型，内容不能为空
        if cmdType.value == 4.0:
            if cmdValue.ctype == 0:
                print('第', i + 1, "行,第2列数据有毛病")
                checkCmd = False
        # 等待类型，内容必须为数字
        if cmdType.value == 5.0:
            if cmdValue.ctype != 2:
                print('第', i + 1, "行,第2列数据有毛病")
                checkCmd = False
        # 滚轮事件，内容必须为数字
        if cmdType.value == 6.0:
            if cmdValue.ctype != 2:
                print('第', i + 1, "行,第2列数据有毛病")
                checkCmd = False
        i += 1
    return checkCmd


# 任务
def mainWork(sheet1, my_time):
    i = 1
    while i < sheet1.nrows:
        # 取本行指令的操作类型
        cmdType = sheet1.row(i)[0]
        if cmdType.value == 1.0:
            # 取图片名称
            img = "./picture/" + sheet1.row(i)[1].value
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(1, "left", img, reTry)
            print("单击左键", img)
        # 2代表双击左键
        elif cmdType.value == 2.0:
            # 取图片名称
            img = "./picture/" + sheet1.row(i)[1].value
            # 取重试次数
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(2, "left", img, reTry)
            print("双击左键", img)
        # 3代表右键
        elif cmdType.value == 3.0:
            # 取图片名称
            img = "./picture/" + sheet1.row(i)[1].value
            # 取重试次数
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(1, "right", img, reTry)
            print("右键", img)
            # 4代表输入
        elif cmdType.value == 4.0:
            inputValue = sheet1.row(i)[1].value
            if sheet1.row(1)[1].value == "yyd.png":
                inputValue = yyd[time_dic[my_time]]
            elif sheet1.row(1)[1].value == "guo.png":
                inputValue = guo[time_dic[my_time]]
            elif sheet1.row(1)[1].value == "yang.png":
                inputValue = yang[time_dic[my_time]]
            pyperclip.copy(inputValue)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)
            print("输入:", inputValue)
            # 5代表等待
        elif cmdType.value == 5.0:
            # 取图片名称
            waitTime = sheet1.row(i)[1].value
            time.sleep(waitTime)
            print("等待", waitTime, "秒")
        # 6代表滚轮
        elif cmdType.value == 6.0:
            # 取图片名称
            scroll = sheet1.row(i)[1].value
            pyautogui.scroll(int(scroll))
            print("滚轮滑动", int(scroll), "距离")
        i += 1


def main(my_key="0", my_time="noon"):
    sheet = []
    file = 'cmd/cmd.xls'
    wb = xlrd.open_workbook(filename=file)
    # if my_time == "tianbao":
    #     sheet.append(wb.sheet_by_index(wb.nsheets-1))
    #     checkCmd = dataCheck(sheet[0])
    # else:
    for i in range(0, wb.nsheets):
        sheet.append(wb.sheet_by_index(i))
    for i in range(0, wb.nsheets):
        checkCmd = dataCheck(sheet[i])
        if not checkCmd:
            break
    print('欢迎使用wcy的自动控制软件~')
    if checkCmd:
        my_key = int(my_key)
        if my_key == -1:
            while True:
                for i in range(0, wb.nsheets):
                    mainWork(sheet[i], my_time)
                    time.sleep(0.1)
                    print("等待0.1秒")
        else:
            for j in range(0, my_key):
                for i in range(0, wb.nsheets):
                    if my_time == "tianbao":
                        mainWork(sheet[wb.nsheets-1], my_time)
                        break
                    mainWork(sheet[i], my_time)
                    time.sleep(0.1)
                    print("等待0.1秒")
    else:
        print('输入有误或者已经退出!')


def Get_NowTime():
    while True:
        now_time = time.strftime("%H:%M").split(":")
        now_time = [int(now_time[0]), int(now_time[1])]
        if now_time == [7, 59]:
            now_time = "morning"
            break
        elif now_time == [11, 40]:
            now_time = "noon"
            break
        elif now_time == [18, 45]:
            now_time = "evening"
            break
        elif now_time == [21, 0]:
            now_time = "night"
            break
        elif now_time == [0, 0]:
            now_time = "tianbao"
            break
    return now_time


if __name__ == "__main__":
    key = input('选择功能: n.做n次 -1.循环到死 \n')
    for num in range(0, int(key)):
        print("等待中\n")
        now_time_str = Get_NowTime()
        main("1", now_time_str)
        print("等待60s")
        time.sleep(60)

# if __name__ == '__main__':
#     file = 'cmd/cmd.xls'
#     # 打开文件
#     wb = xlrd.open_workbook(filename=file)
#     # 通过索引获取表格sheet页
#     sheet1 = wb.sheet_by_index(0)
#     print('欢迎使用不高兴就喝水牌RPA~')
#     # 数据检查
#     checkCmd = dataCheck(sheet1)
#     if checkCmd:
#         key = input('选择功能: 1.做一次 2.循环到死 \n')
#         if key == '1':
#             # 循环拿出每一行指令
#             mainWork(sheet1)
#         elif key == '2':
#             while True:
#                 mainWork(sheet1)
#                 time.sleep(0.1)
#                 print("等待0.1秒")
#     else:
#         print('输入有误或者已经退出!')
