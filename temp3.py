import sys
import pyautogui as pag
import time
from datetime import datetime
import pandas as pd
import sqlite3
import os
import keyboard
import cv2



path ='D:/download/external/'
file_name = 'com_used_log'
table_name = 0
run = 1

previous_mouse_sec = '61'
previous_image_sec = '61'
previous_min = '61'
previous_hour = '25'

pre_x = sys.maxsize
pre_y = sys.maxsize

time_table = {'time': [], 'type': []}
mouse_start = 0
mouse_trigger_time = 2

cam0 = cv2.VideoCapture(0)
# cam1 = cv2.VideoCapture(1)


# time.sleep(5)

print("start!!! external")

while(1):
    current_time = datetime.now()
    year = '{0:04d}'.format(current_time.year)
    month = '{0:02d}'.format(current_time.month)
    day = '{0:02d}'.format(current_time.day)
    hour = '{0:02d}'.format(current_time.hour)
    minute = '{0:02d}'.format(current_time.minute)
    second = '{0:02d}'.format(current_time.second)

    filename = month + day + ' ' + hour + minute + second

    if run == 1:
        if ((int(second) % 2) == 0 and int(second) != int(previous_image_sec)):
            try:
                success, frame = cam0.read()
                if success:
                    cv2.imwrite(path + filename + '.jpg', frame)
            except:
                print(filename, end=': ')
                print('fail frame capture')

            previous_image_sec = second

        try:
            if keyboard.is_pressed('p'):
                print('stop from user', end=': ')
                print(filename)
                run = 0
                time.sleep(1)
        except:
            pass

    else:
        if keyboard.is_pressed('p'):
            print('restart!!!', end=': ')
            print(filename)
            run = 1
            time.sleep(1)