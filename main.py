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
    # path = 'C:\\Kmon20\\'
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

    # a = {'filename': ['tek_0000', 'tek_0001', 'tek_0002', 'tek_0003', 'tek_0004'], 'RUN': [1, 1, 1, 1, 1],
    #  'Auto mode': [0, 0, 0, 0, 0], 'PID-EN': [1, 1, 1, 1, 1], 'V5-EN': [0, 0, 0, 0, 0], 'C12-EN': [0, 0, 0, 0, 0],
    #  'CP-EN': [1, 1, 1, 1, 1], 'RF1-EN': [0, 0, 0, 0, 0], 'RF2-EN': [0, 0, 0, 0, 0], 'SET PID V': [1, 1, 1, 1, 1],
    #  'SET PID C': [1, 1, 1, 1, 1], 'SEL Freq1': [0, 0, 0, 0, 0], 'SEL Freq2': [0, 0, 0, 0, 0],
    #  'SEL Freq3': [0, 0, 0, 0, 0], 'b5': [0, 0, 0, 0, 0], 'b6': [0, 0, 0, 0, 0], 'b7': [0, 0, 0, 0, 0],
    #  '0.Tag Num': ['1', '1', '1', '1', '1'], '1.Switch': [37, 37, 37, 37, 37], '2.Push Button': [3, 3, 3, 3, 3],
    #  '3.CP PWM <= 320': ['10', '10', '10', '10', '10'], '4.CP Volt(V)': ['100', '100', '100', '100', '100'],
    #  '5.PID KP V': ['0.5', '0.5', '0.5', '0.5', '0.5'], '6.PID KI V': ['0.1', '0.1', '0.1', '0.1', '0.1'],
    #  '7.PID KD V': ['0.1', '0.1', '0.1', '0.1', '0.1'], '8.CP Curr(A)': ['0.4', '0.4', '0.4', '0.4', '0.4'],
    #  '9.PID KP C': ['1.5', '1.5', '1.5', '1.5', '1.5'], '10.PID KI C': ['0.3', '0.3', '0.3', '0.3', '0.3'],
    #  '11.PID KD C': ['0.2', '0.2', '0.2', '0.2', '0.2'], '12.RF Freq1(kHz)': ['1000', '1000', '1000', '1000', '1000'],
    #  '13.RF Freq2(kHz)': ['3000', '3000', '3000', '3000', '3000'],
    #  '14.RF Freq3(kHz)': ['20000', '20000', '20000', '20000', '20000'],
    #  '15.Output': ['65535', '65535', '65535', '65535', '65535']}


    test_seq = kmon.make_test_sequence(test_file, path, sheet=test_sheet)
    kmon.save_test_info_initial()
    pag.PAUSE = 0.1
    kmon.test_process(test_seq, kmon_capture)
    kmon.save_test_list(path, info_file_num)