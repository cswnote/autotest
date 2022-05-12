# import pyautogui
# import time
# from datetime import datetime
# # import cv2
# import keyboard
#
# path = 'D:/downloads/'
#
# time.sleep(3)
# run = 1
#
# x1, y1 = pyautogui.position()
# print(x1, y1)
# time.sleep(3)
# x2, y2 = pyautogui.position()
# print(x2, y2)
# time.sleep(3)
# x, y = pyautogui.position()
# print(x, y)
# previous_second = '61'
# previous_min = '61'
#
# width = int(abs(x1-x2))
# height = int(abs(y1-y2))
#
# # cap = cv2.VideoCapture(0)
#
# while(1):
#     current_time = datetime.now()
#     month = '{0:02d}'.format(current_time.month)
#     day = '{0:02d}'.format(current_time.day)
#     hour = '{0:02d}'.format(current_time.hour)
#     minute = '{0:02d}'.format(current_time.minute)
#     second = '{0:02d}'.format(current_time.second)
#     filename = month + day + ' ' + hour + minute + second
#
#     if run == 1:
#
#         if ((int(second) % 30) == 0 and int(previous_second) != int(second)):
#             try:
#                 pyautogui.screenshot(path + filename + '.png', region=(x1, y1, width, height))
#             except:
#                 pass
#
#             previous_second = second
#
#             try:
#                 x2, y2 = pyautogui.position()
#                 pyautogui.rightClick(x, y)
#                 pyautogui.moveTo(x2, y2)
#             except:
#                 pass
#
#         # if (int(minute) % 1) == 0 and int(previous_min) != int(minute) and int(second) == 0:
#         #     try:
#         #         success, frame = cap.read()
#         #         if success:
#         #             cv2.imwrite(path + filename + '.jpg', frame)
#         #     except:
#         #         print(filename, end=': ')
#         #         print('fail frame capture')
#         #
#         #     previous_min = minute
#
#         try:
#             if keyboard.is_pressed("`") and run == 1:
#                 run = 0
#                 print('check pause')
#                 time.sleep(0.5)
#         except:
#             pass
#
#     else:
#         try:
#             if keyboard.is_pressed("`") and run == 0:
#                 run = 1
#                 print('check start')
#                 time.sleep(0.5)
#
#         except:
#             pass


import os
import sys
import pyautogui as pag
import time
from datetime import datetime
import keyboard
# import cv2

# pc_type = 'mac_air'
pc_type = 'home_desktop'

def capture_position_resize(x, y):
    x = int(round(x/1439 * (2560 - 1) * 1.13, 0))
    y = int(round(y/899 * (1600 - 1) * 1.13, 0))
    return x, y

if pc_type == 'mac_air':
    path = '/Users/rainyseason/winston/Workspace/python/Pycharm Project/test/capture/'
elif pc_type == 'home_desktop':
    path = 'D:/downloads/'

run = 1

# x_max, y_max = pag.size()
# if pc_type == 'mac_air':
#     x_max, y_max = position_resize(x_max, y_max)
time.sleep(3)

x1, y1 = pag.position()
if pc_type == 'mac_air':
    x1, y1 = capture_position_resize(x1, y1)
print(x1, y1)
time.sleep(3)

x2, y2 = pag.position()
if pc_type == 'mac_air':
    x2, y2 = capture_position_resize(x2, y2)
print(x2, y2)
time.sleep(3)

x, y = pag.position()
print(x, y)

previous_year = '00'
previous_month = '00'
previous_day = '00'
previous_second = '61'
previous_min = '61'
current_folder = previous_year + previous_month + previous_min + '/'

width = int(abs(x1 - x2))
height = int(abs(y1 - y2))
capture_start_point_x = min(x1, x2)
capture_start_point_y = min(y1, y2)

while True:
    current_time = datetime.now()
    year = format(current_time.year)[2:]
    month = '{0:02d}'.format(current_time.month)
    day = '{0:02d}'.format(current_time.day)
    hour = '{0:02d}'.format(current_time.hour)
    minute = '{0:02d}'.format(current_time.minute)
    second = '{0:02d}'.format(current_time.second)
    filename = month + day + ' ' + hour + minute + second

    if current_folder != year + month + day + '/':
        current_folder = year + month + day + '/'
        try:
            os.mkdir(path + current_folder)
            print('make folder: ', path + current_folder)
        except:
            print('fail make folder: ', path + current_folder)
        # current_folder = current_folder + '/'

    if run == 1:
        if (int(second) % 30) == 0 and int(previous_second) != int(second):
            try:
                pag.screenshot(path + current_folder + filename + '.png', region=(capture_start_point_x, \
                                                                                  capture_start_point_y, width, height))
            except:
                print('fail to save file ', filename)

            previous_second = second

        if (int(second) % 15) == 0:
            try:
                x3, y3 = pag.position()
                if pc_type == 'mac_air':
                    pag.click(x, y)
                elif pc_type == 'home_desktop':
                    pag.rightClick(x, y)
                pag.moveTo(x3, y3)
            except:
                pass

        # if (int(minute) % 1) == 0 and int(previous_min) != int(minute) and int(second) == 0:
        #     try:
        #         success, frame = cap.read()
        #         if success:
        #             cv2.imwrite(path + filename + '.jpg', frame)
        #     except:
        #         print(filename, end=': ')
        #         print('fail frame capture')
        #
        #     previous_min = minute

        try:
            if keyboard.is_pressed("&") and run == 1:
                run = 0
                print('check pause: ', end='')
                print(filename)
                time.sleep(0.5)
        except:
            pass

    else:
        try:
            if keyboard.is_pressed("&") and run == 0:
                run = 1
                print('check start: ', end='')
                print(filename)
                time.sleep(0.5)
        except:
            pass