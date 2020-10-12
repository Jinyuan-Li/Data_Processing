import os
import time
import numpy as np
import pandas as pd
t0 = time.time()

k = 0
year = [2019, 2018, 2017, 2016]
# sum_list = [1,1,1,1,1,1,1]
for i in year:
    # k = 0
    files = os.listdir('\\\\Tspc-nas\\中心資源\\資-數位學習網\\學習網護理資料\\問卷報表\\' + str(i))
    for j in files:
        path = '\\\\Tspc-nas\\中心資源\\資-數位學習網\\學習網護理資料\\問卷報表\\' + str(i) + '\\' + j
        j = j.replace('.xlsx', '')
        df1 = pd.read_excel(path, sheet_name='問卷統計')
        df1.to_excel('\\\\Tspc-nas\\中心資源\\資-數位學習網\\學習網護理資料\\問卷報表\\xlsx彙整\\' + j + '.xlsx', sheet_name='問卷統計', index=0, header=0)

        df2 = pd.read_excel(path, sheet_name='答題記錄')
        df2.to_excel('\\\\Tspc-nas\\中心資源\\資-數位學習網\\學習網護理資料\\問卷報表\\xlsx原始\\' + j + '.xlsx', sheet_name='答題記錄', index=0, header=0)

        # df2 = df.iloc[1:30, 0:5]
        # df2[6] = j.replace('.xlsx', '')
        # df2[7] = i
        # df2 = df2[[6, 7, '敬請填寫課後評估表 - 統計', 'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4']]
        # sum_list = np.vstack((sum_list, df2))
        print(k)
    #     k += 1
    #     if k == 1:
    #         break
    # if k == 1:
    #     break

# df_output = pd.DataFrame(sum_list)
# df_output.to_excel(r'C:\Users\user\Desktop\test.xlsx', sheet_name='test', index=0, header=0)
print('花費時間：',time.time()-t0)