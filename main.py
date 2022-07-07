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
        filelist = [int(file.split('tek')[1][:4]) for file in filelist if file[:3] == 'tek']
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
    test_sheet = 'manual mode'
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

