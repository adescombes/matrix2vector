import os
import pandas as pd
import sys
sys.path.insert(1, r'C:\Users\abe\OneDrive - TRANSITEC\Documents - 2314_230-CH-EBBN-BielWestGesamtmobilitat\General\3-Ingenierie\2-Skripte')
from functions import config_1, config_2

writer = pd.ExcelWriter('test.xlsx', engine = 'xlsxwriter')
excel_file = pd.ExcelFile('GVM BE 2019_MIV_DTV_Ist 2019_Q-Z-Matrix.xlsx')
sheet_names_list = excel_file.sheet_names

for s_n in sheet_names_list:

    print(s_n)
    try:
        df_vector = config_1(excel_file, s_n, writer)
        df_vector.to_excel(writer, index = False, sheet_name = s_n)
    except Exception as e:
        print(e)

writer.close()
        

