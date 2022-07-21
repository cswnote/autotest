# import os
# import sys
# import pyautogui as pag
# import time
# from datetime import datetime
# import keyboard
#
# while True:
#     x, y = pag.position()
#     print(x, y)
#     time.sleep(2)


import os
import sys
import shutil

path = '/Users/rainyseason/winston/Workspace/python/Pycharm Project/sinewave_analyze/Evaluation/tek_csv/'

files = os.listdir(path)

files = [file for file in files if file[:3] == 'tek']

files.sort()

for file in files:
    num = int(file[3:-4])
    if 8786 <= num <= 8962:
        num  = num - 13
        src = path + file
        extention = file[-3:]
        file = file[0:4] + str(num) + '.' + extention
        if extention == 'png':
            dst = '/Users/rainyseason/winston/Workspace/python/Pycharm Project/sinewave_analyze/Evaluation/tek_png/' + file
        elif extention == 'csv':
            dst = '/Users/rainyseason/winston/Workspace/python/Pycharm Project/sinewave_analyze/Evaluation/tek_csv/' + file
        else:
            dst = '/Users/rainyseason/winston/Workspace/python/Pycharm Project/sinewave_analyze/Evaluation/tek_set/' + file

        shutil.copy(src, dst)
