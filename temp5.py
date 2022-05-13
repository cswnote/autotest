import os
import sys
import pyautogui as pag
import time
from datetime import datetime
import keyboard

path = os.getcwd() #+ '\\PL150_WS_eval\\'
filelist = os.listdir(path)
a=[]
for file in filelist:
    if file[:3] == 'tek':
        a.append(file.split('tek')[1][:4])


filelist = [int(file.split('tek')[1].split('_kmonCap')[0]) for file in filelist if file[:3] == 'tek']

for file in filelist:
    src = os.path.join(path, file + '.png')
    dst = file + '_kmonCap.png'
    dst = os.path.join(path, dst)
    os.rename(src, dst)

