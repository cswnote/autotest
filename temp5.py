import os
import sys
import pyautogui as pag
import time
from datetime import datetime
import keyboard

path = os.getcwd() #+ '\\PL150_WS_eval\\'
path = path[:-13]
filelist = os.listdir(path)

filelist = [file.split('.')[0] for file in filelist if file[:3] == 'tek']

for file in filelist:
    src = os.path.join(path, file + '.png')
    dst = file + '_kmonCap.png'
    dst = os.path.join(path, dst)
    os.rename(src, dst)

