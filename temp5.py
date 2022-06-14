import os
import sys
import pyautogui as pag
import time
from datetime import datetime
import keyboard

while True:
    x, y = pag.position()
    print(x, y)
    time.sleep(2)
