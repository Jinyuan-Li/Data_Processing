import os
import time
import numpy as np
import pandas as pd
t0 = time.time()

sum_list = []
files = os.listdir(r'C:\Users\user\Desktop\elearning_exam2020fin')

for n, i in enumerate(files):
    tem_list = [i.replace('.xlsx', '')]
    df = pd.read_excel(r'C:\\Users\\user\\Desktop\\elearning_exam2020fin\\' + i, sheet_name='測驗統計')
    df = df.iloc[1:, :]
    for j in range(df.shape[0]):
        if len(df.iloc[j, 0]) > 5:
            tem_list.append(df.iloc[j, 0])
            for k in range(2):
                tem_list.append('--')
        tem_list.append(df.iloc[j, 1])
        tem_list.append(df.iloc[j, 3])
        if df.iloc[j, 1] == '否':
            for k in range(4):
                tem_list.append('--')
    sum_list.append(tem_list)

print('花費時間：',time.time()-t0)
df_output = pd.DataFrame(sum_list)
df_output.to_excel(r'C:\Users\user\Desktop\elearning_exam2020-2.xlsx', sheet_name='elearning_exam', index=0, header=0)