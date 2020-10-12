import time
import numpy as np
import pandas as pd
t0 = time.time()

df = pd.read_excel(r'D:\OneDrive\文件\Analysis\APP BSRS檢測及開啟紀錄\心情溫度計(App)開啟紀錄20200810摘要 - 開啟sheet.xlsx', sheet_name='開啟')
county_list = ['新北市', '台北市', '桃園市', '台中市', '台南市', '高雄市', '宜蘭縣', '新竹縣', '苗栗縣', '彰化縣', '南投縣',
               '雲林縣', '嘉義縣', '屏東縣', '台東縣', '花蓮縣', '澎湖縣', '基隆市', '新竹市', '嘉義市', '金門縣', '連江縣', '外國'] # 欽榮順序
yymm_list = ['2019-04', '2019-05', '2019-06', '2019-07', '2019-08', '2019-09', '2019-10', '2019-11', '2019-12',
             '2020-01', '2020-02', '2020-03', '2020-04', '2020-05', '2020-06', '2020-07']

df_output = pd.DataFrame(np.zeros((23, len(yymm_list))), index=county_list, columns=yymm_list)

set_uid = []
for i in range(df.shape[0]):
    if isinstance(df.iloc[i, 4], str) == 1: # 城市非空白
        if df.iloc[i, 1] not in set_uid: # UID是否首次出現
            set_uid.append(df.iloc[i, 1])
            df_output.loc[df.iloc[i, 4], df.iloc[i, 5][0:7]] += 1

df_output.to_excel(r'C:\Users\user\Desktop\test.xlsx', sheet_name='test', index=1, header=1)
print('花費時間：',time.time()-t0)
