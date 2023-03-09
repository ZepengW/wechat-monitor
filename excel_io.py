import pandas as pd
import os
import logging


def search_aid(excel_path):
    aid_dict = dict()
    if not os.path.isfile(excel_path):
        logging.debug(f'file not exist: {excel_path}')
        return aid_dict
    f = pd.ExcelFile(excel_path)
    for sheetname in f.sheet_names:
        d = pd.read_excel(excel_path, sheet_name = sheetname)
        aid_dict[sheetname] = list(d['aid'])
    return aid_dict

def add_new_item(excel_path, data_list):
    df = pd.DataFrame(data_list)


def write_csv(path, data_list, mode = 'w'):
    df = pd.DataFrame(data_list)
    if mode == 'w':
        df.to_csv(path, ',', index=False)
    elif mode == 'a':
        df_old = pd.read_csv(path)
        df = pd.concat([df, df_old])
        df = df.drop_duplicates()
        df = df.sort_values('time')
        df.to_csv(path, ',', index=False)

if __name__ == '__main__':
    search_aid('/Users/wangzepeng/Downloads/test.xlsx')