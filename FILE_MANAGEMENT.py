import os
import sqlite3
import pandas as pd
import numpy as np
import gc
import openpyxl

class FILE_MANAGEMENT():
    def __init__(self):
        super().__init__()


    def create_folder(self, path, directory):
        try:
            if not os.path.exists(path + directory):
                os.makedirs(path + directory)
                # print("creat '%s' folder" % directory)
        except OSError:
            print('Error: Creating directory. ' + path)

        return (path + directory + '/')

    def find_specific_extention_file(self, extention, path):
        file_list = os.listdir(path)
        file_list = [file for file in file_list if file.endswith(extention)]

        for i in range(len(file_list)):
            file_list[i] = file_list[i].split(extention)[0]

        return file_list


    def db_to_csv(self, db_file_name, db_path, csv_file_name, csv_path, table_name, default_index_name):
        # print('61')
        print(db_path + db_file_name + '.db')
        con = sqlite3.connect(db_path + db_file_name + '.db')
        # print('62')
        print(table_name)
        df = pd.read_sql("SELECT * FROM [%s]" % table_name, con, index_col=default_index_name)
        # print('63')
        df.to_csv(csv_path + csv_file_name + ' - ' + db_file_name + '.csv', mode='w', encoding='utf-8-sig')

        print('success to create csv from dataframe')


    def db_to_excel(self, table_name, db_file_name, db_path, excel_file_name, excel_path):
        con = sqlite3.connect(db_path + db_file_name + '.db')
        df = pd.read_sql("SELECT * FROM [%s]" % table_name, con, index_col='index')
        df.to_excel(excel_path + excel_file_name + '.xlsx')
        del df
        gc.collect()
        print('done db to excel: %s' % db_file_name)


    def data_to_file(self, dict_data, file_name, path):
        main_data = dict_data.copy()
        key_list = list(main_data.keys())
        col_name = []
        row_data = {}

        print(path + file_name)
        con = sqlite3.connect(path + file_name + '.db')
        cursor = con.cursor()

        for i in range(len(key_list) - 4, 4, -1):
            del main_data[key_list[i]]

        self.add_db_table_from_dictdata(main_data, '0000_main', con)

        for i in range(0, 5):
            del dict_data[key_list[i]]

        for i in range(len(dict_data['기간'])):
            col_name.append(dict_data['기간'][i])
            if len(col_name) != 0:
                col_name = col_name[0]
                break

        for i in range(len(col_name)):
            for j in range(len(col_name) - 1, i, -1):
                if col_name[i] == col_name[j]:
                    col_name[i] = col_name[i] + 'y'
                    break

        for i in range(len(key_list) - 1, len(key_list) - 5, -1):
            del dict_data[key_list[i]]

        key_list = list(dict_data.keys())

        for s_idx in range(len(main_data['index'])):
            for i, key in enumerate(key_list):
                row_data[key] = []
            for row_idx, key in enumerate(key_list):
                row_data[key] = dict_data[key][s_idx]

            df = pd.DataFrame.from_dict(row_data, orient='index', columns=col_name)
            table_name = str(format(main_data['index'][s_idx] + 1, '04')) + '-' + str(main_data['code'][s_idx]) + '-' + main_data['name'][s_idx]
            df.to_sql(table_name, con, if_exists='replace')
            print('success to add %s table to DB' % table_name)
            del df

        con.close()
        gc.collect()


    def add_db_table_from_dictdata(self, dict_data, table_name, connect):
        df = pd.DataFrame.from_dict(dict_data, orient='index')
        df = df.transpose()

        df.to_sql(table_name, connect, if_exists='replace', index=False)

        del df
        gc.collect()


    def add_average_to_dataframe(self, dataframe, column_name):
        # # get average of 'column_name'
        average = dataframe[column_name].mean()
        print(average)
        # dataframe = dataframe.append({'buy_date': 'average', 'buy_price': np.nan, 'gain': average, 'holding_period': np.nan}, ignore_index=True)
        dataframe = dataframe.append({'buy_date': 'average', 'buy_price': np.nan, 'flag': np.nan, 'tr_date': np.nan, 'sell_price': np.nan, 'nominal_gain': np.nan, 'real_gain': np.nan, 's_rate': np.nan, 'gain_sum': np.nan, 'holdings': np.nan, 'gain_per_tr': average}, ignore_index=True)

        return dataframe


    def make_db_file_from_datafram(self, dataframe, table_name, connect):
        # make pandas dataframe from dictionary data
        df = dataframe
        df.to_sql(table_name, connect, if_exists='replace')

        connect.commit()
        connect.close()

        del df
        gc.collect()

        print('success to change db from dataframe')


    def change_str_from_db(self, target_string, to_be_string, db_name, path):
        con = sqlite3.connect(path + db_name + '.db')
        cursor = con.cursor()
        cursor.execute(
            "SELECT name FROM sqlite_master WHERE type IN('table', 'view') AND name NOT LIKE 'sqlite_%' UNION ALL SELECT name FROM sqlite_temp_master WHERE type IN ('table', 'view') ORDER BY 1")
        table_list = cursor.fetchall()

        size = len(table_list)
        for i in range(size):
            print('in process: %i/%i : %s' % (i + 1, len(table_list), table_list[i][0]))
            table_list[i] = table_list[i][0]
            if table_list[i].find(target_string) != -1:
                temp_table = table_list[i].replace(target_string, to_be_string)
                cursor.execute("ALTER TABLE [%s] RENAME TO [%s]" % (table_list[i], temp_table))
        con.commit()
        con.close()


    def get_table_list_from_cursor(self, cursor):
        cursor.execute(
            "SELECT name FROM sqlite_master WHERE type IN('table', 'view') AND name NOT LIKE 'sqlite_%' UNION ALL SELECT name FROM sqlite_temp_master WHERE type IN ('table', 'view') ORDER BY 1")
        table_list = cursor.fetchall()

        for i in range(len(table_list)):
            table_list[i] = table_list[i][0]

        return table_list


    def db_data_integrity_check(self, check_string, file, path):
        con = sqlite3.connect(path + file + '.db')
        cursor = con.cursor()
        table_list = self.get_table_list_from_cursor(cursor)
        error_list = {'num': [], 'name': []}
        blank_list = {'num': [], 'name': []}

        for i in range(len(table_list)):
            print('in process: %i/%i : %s' % (i + 1, len(table_list), table_list[i]))
            df = pd.read_sql("SELECT * FROM [%s]" % table_list[i], con, index_col='index')
            if len(df) > 0:
                if df.axes[0][0] != check_string:
                    error_list['name'].append(table_list[i])
                    error_list['num'].append(i)
            else:
                blank_list['name'].append(table_list[i])
                blank_list['num'].append(i)


        con.commit()
        con.close()

        return error_list, blank_list


    def make_db_file_from_dictdata(self, dict_data, table_name, file_name, db_path):
        # make pandas dataframe from dictionary data
        df = pd.DataFrame.from_dict(dict_data, orient='index')
        df = df.transpose()
        # # open db file
        con = sqlite3.connect(db_path + file_name + '.db')

        df = self.add_average_to_dataframe(df, 'gain')

        df.to_sql(table_name, con, if_exists='replace')

        con.commit()
        con.close()

        del df
        gc.collect()


if __name__ == '__main__':
    print('in FILE_MANAGEMENT')