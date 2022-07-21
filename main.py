import pyautogui as pag
import pyperclip as clip
import time
import KMON
import os
from FILE_MANAGEMENT import *
import TEK_SCOPE
import openpyxl


if __name__ == '__main__':
    path = os.getcwd() + '/PL150_WS_eval/'
    kmon_capture = True
    # kmon_capture = False
    tek_serial_num = 'C040861'
    filelist = os.listdir(path)
    try:
        filelist = [int(file.split('tek')[1][:5]) for file in filelist if file[:3] == 'tek']
        filelist = max(filelist)
    except:
        filelist = -1
    start_file_num = filelist + 1
    filelist = os.listdir(path)
    try:
        filelist = [int(file.split('info_test_')[1].split('.')[0]) for file in filelist if file[:10] == 'info_test_']
        filelist = max(filelist)
    except:

        filelist = -1

    test_file = 'test.xlsx'
    test_sheet = 'manual mode ch4'
    info_file_num = filelist + 1
    del filelist

    # tek_serial_num = input("type your scope's serial number: ")
    # file_num = 0
    # filename = 'tek_' + ('{0:05d}'.format(file_num))

    time.sleep(1)
    pag.PAUSE = 0.05

    sw = {}
    pb = {}
    tx_packets = {}
    test_seq = []
    loop_count = 0
    inner_loop = []
    inner_loop_count = 0

    kmon = KMON.KMON(start_file_num, info_file_num, path)

    # # initial scope
    kmon.set_default_scope(tek_serial_num)
    kmon.scope_run_single_mode('single')
    # kmon.scope_run_single_mode('continuous')
    kmon.scope_on()
    kmon.check_filesystem()
    kmon.save_file_format(image_type='png', ink_save='on', savewave_type='csv')

    kmon.origin_coordinate()
    sw, pb, tx_packets = kmon.get_packet_info()
    pag.PAUSE = 0.1

    test_seq = kmon.make_test_sequence(test_file, path, sheet=test_sheet)
    kmon.save_test_info_initial()
    pag.PAUSE = 0.1
    pag.moveTo(257, 169)
    pag.click(button='right')
    kmon.test_process(test_seq, kmon_capture)
    pag.moveTo(257, 169)
    pag.click(button='right')

    kmon.save_test_list(path, info_file_num)
    path = path + 'kmon csv/'
    kmon.save_kmon_csv(path, info_file_num)





    # time.sleep(1)
    # info_file_num = 1
    # path = 'D:/winston/workspace/Pycharm Projects/autotest/PL150_WS_eval/kmon csv'
    #
    #
    # pag.hotkey('Alt', 'tab')
    # pag.moveTo(1382, 116)
    # pag.click(button='left')
    # time.sleep(0.5)
    # pag.move(88, 195)
    # time.sleep(0.5)
    # pag.move(164, 23)
    # time.sleep(0.5)
    # pag.click(button='right')
    # time.sleep(5)
    #
    # pag.keyDown('alt')
    # pag.press('tab', presses=2, interval=0.5)
    # pag.keyUp('alt')
    #
    # time.sleep(1)
    # pag.press('Alt')
    # time.sleep(1)
    # pag.press('f')
    # time.sleep(1)
    # pag.press('a')
    # time.sleep(1)
    # pag.press('o')
    # time.sleep(1)
    # pag.moveTo(1914, 37)
    # pag.hotkey('Alt', 'd')
    # pag.write(path)
    # pag.press('enter')
    # time.sleep(1)
    # pag.hotkey('alt', 't')
    # pag.press('down', presses=5, interval=0.5)
    # pag.press('enter')
    # time.sleep(1)
    # pag.hotkey('Alt', 'n')
    # pag.write('info_kmon_' + ('{0:02d}'.format(info_file_num)) + '_1.csv')
    # pag.hotkey('alt', 's')
    # time.sleep(5)
    # pag.hotkey('alt', 'f4')
    #
    # time.sleep(1)
    # pag.moveTo(1400, 590)
    # pag.click(button='left')
    # time.sleep(0.5)
    # pag.move(88, 195)
    # time.sleep(0.5)
    # pag.move(164, 23)
    # time.sleep(0.5)
    # pag.click(button='right')
    # time.sleep(5)
    #
    # pag.keyDown('alt')
    # pag.press('tab', presses=2, interval=0.5)
    # pag.keyUp('alt')
    #
    # time.sleep(1)
    # pag.press('Alt')
    # time.sleep(1)
    # pag.press('f')
    # time.sleep(1)
    # pag.press('a')
    # time.sleep(1)
    # pag.press('o')
    # time.sleep(1)
    # pag.moveTo(1914, 37)
    # pag.hotkey('Alt', 'd')
    # pag.write(path)
    # pag.press('enter')
    # time.sleep(1)
    # pag.hotkey('alt', 't')
    # pag.press('down', presses=5, interval=0.5)
    # pag.press('enter')
    # time.sleep(1)
    # pag.hotkey('Alt', 'n')
    # pag.write('info_kmon_' + ('{0:02d}'.format(info_file_num)) + '_2.csv')
    # pag.hotkey('alt', 's')
    # time.sleep(5)
    # pag.hotkey('alt', 'f4')
