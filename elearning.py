import os
import time
import numpy as np
import pandas as pd
t0 = time.time()

year = [2019, 2018, 2017, 2016]
sum_list = [['課程名稱', '年份', '通過人次',  '修課人次', '修課ID總累計']]
ID_list_T = [[]] # ID總累計
# take_o1_list = []
# take_o2_list = []
date_list = []
df_date = pd.DataFrame(np.zeros((12, 6)), index = [1, 2, 3, 4, 5, 6, 7, 8, 9 ,10, 11, 12], columns=[2015, 2016, 2017, 2018, 2019, 2020])

for i in year:
    ID_list = [] # ID年累計
    z = 0

    files = os.listdir('\\\\Tspc-nas\\中心資源\\資-數位學習網\\學習網護理資料\\課程通過報表\\' + str(i))
    for j in files:
        tem_list = []
        tem_list.append(j.replace('_member.xlsx', ''))
        tem_list.append(i)
        df = pd.read_excel('\\\\Tspc-nas\\中心資源\\資-數位學習網\\學習網護理資料\\課程通過報表\\' + str(i) + '\\' + j, sheet_name='member')

        # 通過人次
        df_pass = df.iloc[1:, 7]
        tem_list.append(df_pass[df.iloc[:, 7] == '通過'].count())

        # 修課人次、修課ID累計
        df_take = df.iloc[1:, 1]
        tem_list.append(len(df_take))
        ID_list = np.hstack((ID_list, df_take))
        ID_list_T[0] = np.hstack((ID_list_T[0], df_take))
        tem_list.append(len(set(ID_list)))

        # # 通過人次、修課人次分男女
        # pass_m, pass_f, take_m, take_f = 0, 0, 0, 0
        # take_o1, take_o2 = 0, 0
        #
        # for p, q in zip(df_take, df_pass):
        #     if type(p) != float:
        #         if p[1] == '1' or p[1] == 'A' or p[1] == 'C' or p[1] == 'a' or p[1] == 'c' or p[1] == '0': # https://www.immigration.gov.tw/5385/7244/7250/20406/7326/36526/
        #             take_m += 1
        #             if q == '通過':
        #                 pass_m += 1
        #         elif p[1] == '2' or p[1] == 'B' or p[1] == 'D' or p[1] == 'b' or p[1] == 'd' or p[1] == 'G': # 0、G為例外個案
        #             take_f += 1
        #             if q == '通過':
        #                 pass_f += 1
        #     #     else:
        #     #         take_o1 += 1
        #     #         take_o1_list.append(p)
        #     # else:
        #     #     take_o2 += 1
        #     #     take_o2_list.append(p)
        #
        # tem_list.append(take_m)
        # tem_list.append(take_f)
        # tem_list.append(pass_m)
        # tem_list.append(pass_f)
        # tem_list.append(take_o1)
        # tem_list.append(take_o2)

        # 完成日期
        date = df.iloc[1:, 10]
        for k in date:
            if k!=  '-':
                yy = int(k.split('-')[0])-2015
                mm = int(k.split('-')[1])-1
                df_date.iloc[mm, yy] += 1

        print(tem_list)
        sum_list.append(tem_list)

        # if z == 1:
        #     break
        # z += 1

#     ID_list_T.append(set(ID_list))
# ID_list_T[0] = set(ID_list_T[0])

# sum_list[0].append('修課人次(男)')
# sum_list[0].append('修課人次(女)')
# sum_list[0].append('通過人次(男)')
# sum_list[0].append('通過人次(女)')
# sum_list[0].append('修課人次(例外)')
# sum_list[0].append('修課人次(空白)')

# print(sum_list)
df_output = pd.DataFrame(sum_list)
df_output.to_excel(r'C:\Users\user\Desktop\test.xlsx', sheet_name='test', index=0, header=0)

# df_output = pd.DataFrame(take_o1_list)
# df_output.to_excel(r'C:\Users\user\Desktop\修課人次(例外).xlsx', sheet_name='test', index=0, header=0)
#
# df_output = pd.DataFrame(take_o2_list)
# df_output.to_excel(r'C:\Users\user\Desktop\修課人次(空白).xlsx', sheet_name='test', index=0, header=0)

df_date.to_excel(r'C:\Users\user\Desktop\date_list.xlsx', sheet_name='test', index=0, header=0)

# df_output = pd.DataFrame(ID_list_T)
# df_output = df_output.transpose()
# df_output.to_excel(r'C:\Users\user\Desktop\ID_list_T.xlsx', sheet_name='test', index=0, header=0)

print('花費時間：',time.time()-t0)