import csv
import pandas as pd

def write_csv(path, data_list, mode = 'w'):
    df = pd.DataFrame(data_list)
    if mode == 'w':
        df.to_csv(path, mode, index=False)
    elif mode == 'a':
        df_old = pd.read_csv(path)
        df = pd.concat([df, df_old])
        df = df.drop_duplicates()
        df = df.sort_values('time')
        df.to_csv(path, 'w', index=False)