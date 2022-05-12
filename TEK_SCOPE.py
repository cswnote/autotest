import pyvisa
from datetime import datetime
import sys
import time


class TekSCOPE():
    def __init__(self):
        super().__init__()

        self.rm = 0

        self.scope = 0
        self.scope_mode = ''

        self.image_format = 'png'
        self.waveform_format = 'csv'

        self.system_path = 'E:/'

        self.csv_save_delay = 0
        self.previous_time = 0

        self.file_num = 0

    def set_default_scope(self, serial_num):
        self.rm = pyvisa.ResourceManager()
        instrument_list = self.rm.list_resources()

        for inst in instrument_list:
            try:
                if inst.split('::')[3] == serial_num:
                    self.scope = self.rm.open_resource(inst)
            except:
                pass
        print('information of oscilloscope: ', end=' ')
        print(self.scope.query('*IDN?'))

        horizontal_info = self.scope.query('HORizontal?')
        record_length = horizontal_info.split(';')[3]

        if record_length == '1000':
            self.csv_save_delay = 2
        elif record_length == '10000':
            self.csv_save_delay = 2.4
        elif record_length == '100000':
            self.csv_save_delay = 3.5
        elif record_length == '1000000':
            self.csv_save_delay = 12        # 이하 검증 안됨
        elif record_length == '5000000':
            self.csv_save_delay = 54
        elif record_length == '10000000':
            self.csv_save_delay = 105
        else:
            print('scope error detect')
            sys.exit()


    def save_file_format(self, **kwargs):
        self.image_format = kwargs.get('image_type', 'PNG')
        image_ink_save = kwargs.get('ink_save', 'ON')
        self.waveform_format = kwargs.get('wave_type', 'SPREADSheet')

        if self.waveform_format == 'csv' or self.waveform_format == 'xlsx' or self.waveform_format == 'xls' or \
                self.waveform_format == 'spreadsheet' or self.waveform_format == 'SPREADSheet':
            self.waveform_format = 'SPREADSheet'

        # if image_format == 'png':
        #     image_format == 'PNG'
        #
        # if image_format == 'bmp':
        #     image_format == 'BMP'
        self.image_format = self.image_format.upper()

        now = datetime.now()
        while(1):
            self.scope.write('SAVe:IMAGe:FILEFormat ' + self.image_format)
            if self.scope.query('SAVe:IMAGe:FILEFormat?') == self.image_format + '\n':
                print("image file format is " + self.image_format + '.')
                break
            if (datetime.now() - now).seconds > 10:
                print('error: scope is not answered or check image format')
                sys.exit()

        now = datetime.now()
        while(1):
            self.scope.write('SAVe:WAVEform:FILEFormat ' + self.waveform_format)
            if self.scope.query('SAVe:WAVEform:FILEFormat?') == 'SPREADSHEET\n':
                print("waveform file format is 'csv'")
                break
            if (datetime.now() - now).seconds > 10:
                print('error: scope is not answered')
                sys.exit()

        self.image_format = self.image_format.lower()
        if self.waveform_format == 'SPREADSheet':
            self.waveform_format = 'csv'


    def save_files(self, filename):
        print("SAVe:SETUp '" + filename + ".SET'")
        self.scope.write("SAVe:SETUp '" + filename + ".SET'")
        self.scope.write("SAVe:IMAGe '" + filename + '.' + self.image_format + "'")
        self.scope.write("SAVE:WAVEFORM ALL, '" + filename + '.' + self.waveform_format + "'")
        time.sleep(self.csv_save_delay)


    def check_filesystem(self):
        self.scope.write("FILESystem:CWD 'E/'")
        print(self.scope.query('FILESystem?'))
        print('make later')


    def scope_run_single_mode(self, mode='single'):
        now = datetime.now()
        if mode == 'continuous':
            while(1):
                self.scope.write('ACQuire:STOPAfter RUNSTop')
                if self.scope.query('ACQuire:STOPAfter?') == 'RUNSTOP\n':
                    self.scope_mode = 'continuous'
                    print("Oscilloscope mode is Continuous mode.")
                    break
                if (datetime.now() - now).seconds > 10:
                    print('error: scope is not answered.')
                    sys.exit()

        if mode == 'single':
            while(1):
                self.scope.write('ACQuire:STOPAfter SEQuence')
                if self.scope.query('ACQuire:STOPAfter?') == 'SEQUENCE\n':
                    self.scope_mode = 'single'
                    print("Oscilloscope mode is Single mode.")
                    break
                if (datetime.now() - now).seconds > 10:
                    print('error: scope is not answered.')
                    sys.exit()




if __name__ == '__main__':
    my_tek_serial_num = 'C021083'
    file_num = 0
    filename = 'tek_' + ('{0:04d}'.format(file_num))

    Tek = TekSCOPE()
    Tek.set_default_scope(my_tek_serial_num)
    Tek.scope_run_single_mode('single')
    Tek.check_filesystem()
    Tek.save_file_format(image_type='png', image_ink_save='on', savewave_type='csv')
    for i in range(0, 100):
        if Tek.scope_mode == 'single':
            Tek.scope.write('ACQuire:STATE ON')
        elif Tek.scope_mode == 'continuous':
            Tek.scope.write('ACQuire:STATE ON')
            Tek.scope.write('ACQuire:STATE OFF')
            # pass
        Tek.save_files(filename)
        print(filename)
        file_num += 1
        filename = 'tek_' + ('{0:04d}'.format(file_num))
