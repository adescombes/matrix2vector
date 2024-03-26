import os
import pandas as pd

CHUNK_SIZE = 1000000 # max sheet size is 1048576

"""
CONFIG 1

| -------- | -------- | 1        | 1        | 1
| -------- | NAME     | Dest 1   | Dest 1   | Dest 1
| MATRIX   | -------- | Nr 5     | Nr 6     | Nr 7
| idx 1    | Orig 1   | 0        | 0        | ...
| idx 2    | Orig 2   | 0        | 0        | ...
| idx 3    | Orig 3   | 0        | 0        | ...

"""
def config_1(excel_file, s_n, writer):

    print('config_1')
    df = pd.read_excel(excel_file, header = [1,2], index_col=[0,1], sheet_name = s_n)

    unnamed_cols = [(x, y) for (x, y) in df.columns if "Unnamed" in y]
    df.drop(unnamed_cols, axis = 1, inplace = True)
    df1 = df.stack(level = 0).reset_index()
    df1.rename(columns = { 'level_0' : 'origin_index', 'level_1' : 'origin_name', 'Name' : 'destination_name' }, inplace = True)

    df1[~df1['origin_name'].isna()]
    dict_idx = dict(zip(df1['origin_name'], df1['origin_index']))
    df1['destination_index'] = df1['destination_name'].map(dict_idx)

    nr_cols = [x for x in df1.columns if "Nr." in x]
    df1['total'] = df1[nr_cols].sum(axis=1)
    print(df1)
    df1 = df1[df1['total'] != 0]
    df1 = df1[~df1['origin_name'].isna()]
    df1 = df1[['origin_index', 'origin_name', 'destination_index', 'destination_name', 'total']].sort_values(by = ['origin_index', 'destination_index'])

    if len(df1) > CHUNK_SIZE:
        print("Size is too large, writing chunks (%d rows)" % len(df1))
        write_xlsx_chunks(df1, s_n, writer)
    else:
        print("Writing to excel...")
        df1.to_excel(writer, index = False, sheet_name = s_n)

    return 

"""
CONFIG 2

| -------- | -------- | -------- | idx 1    | idx 2    |
| -------- | -------- |VON / NACH| Dest 1   | Dest 2   |
| -------- | idx 1    | Orig 1   | 0        | 0        | ...
| -------- | idx 2    | Orig 2   | 0        | 0        | ...
| -------- | idx 3    | Orig 3   | 0        | 0        | ...
| -------- | idx 4    | Orig 4   | 0        | 0        | ... 

"""
def config_2(excel_file, s_n, writer):

    print('config_2')
    df = pd.read_excel(excel_file, sheet_name = s_n, header = 1, index_col=[1, 2])
    df.dropna(how = 'all', axis = 1, inplace = True)
    df2 = df.stack(level = 0).reset_index()
    print(df2.head(10))
    print(df2.columns)
    dict_col = dict(zip(df2.columns, ['origin_index', 'origin_name', 'destination_name', 'total']))
    df2.rename(columns = dict_col, inplace = True)
    dict_idx = dict(zip(df2['origin_name'], df2['origin_index']))
    df2['destination_index'] = df2['destination_name'].map(dict_idx)
    df2 = df2[df2['total'] != 0]
    df2 = df2[~df2['origin_name'].isna()]
    #df3 = df3[['origin_index', 'origin_name', 'destination_index', 'destination_name', 'total']].sort_values(by = ['origin_index', 'destination_index'])
    
    if len(df2) > CHUNK_SIZE:
        print("Size is too large, writing chunks (%d rows)" % len(df2))
        write_xlsx_chunks(df2, s_n, writer)
    else:
        print("Writing to excel...")
        df2.to_excel(writer, index = False, sheet_name = s_n)
    
    return 


def write_xlsx_chunks(df, sheet_name, writer):

    i = 1
    row_start = 0
    row_end = CHUNK_SIZE
    remaining_rows = len(df)

    while remaining_rows > 0:

        print('Writing chunk %d...' % i)
        sheet_name_i = "%s_%d" % (sheet_name, i)
        last_row = min(len(df), row_end)
        df_chunk = df.iloc[ row_start : last_row ] # -> 1000000 rows
        row_start = row_end
        row_end += CHUNK_SIZE
        i += 1
        remaining_rows -= CHUNK_SIZE
        print(df_chunk)
        df_chunk.to_excel(writer, index = False, sheet_name = sheet_name_i)

    return





