import time
import numpy as np
import pandas as pd
from collections import Counter
t0 = time.time()

year_list = [2016, 2017, 2018, 2019]
suggestion_sum_list = []
df_wordcount = pd.DataFrame(np.zeros((10001, 5)), columns=[2016, 2017, 2018, 2019, 'Total'])
with open('stop_words.txt', encoding="utf-8") as f:
    stop_list = f.read().split('\n')


for yr in year_list:
    # if yr == 2016:
        yr_cnt = Counter()
        suggestion_list1 = []
        locals()["df_%s" % str(yr)] = pd.read_excel(r'C:\Users\user\Desktop\suggestions.xlsx', sheet_name=str(yr))
        for row in range(locals()["df_%s" % str(yr)].shape[0]):
            suggestion = locals()["df_%s" % str(yr)].iloc[row, 0]

            # 製作yr x 字數之交叉表
            if len(suggestion) < 9999:
                yr_cnt[len(suggestion)] += 1
            else:
                yr_cnt[10000] += 1

            # 輸出20字以上之留言
            suggestion = suggestion.replace('\n', '')
            for word in stop_list:
                suggestion = suggestion.replace(word, '')
            if len(suggestion) >= 20:
                suggestion_list1.append(suggestion)

        tem_df = pd.DataFrame(yr_cnt.most_common())
        for row_cnt in range(tem_df.shape[0]):
            df_wordcount.loc[tem_df.iloc[row_cnt, 0], [yr, 'Total']] += tem_df.iloc[row_cnt, 1]

        suggestion_list1_set = set(suggestion_list1)
        for suggestionr in suggestion_list1_set:
            suggestion_sum_list.append([len(suggestionr), suggestionr])
        print(yr)
        print('花費時間：%.2f秒' % (time.time()-t0))

df_wordcount.to_excel(r'C:\Users\user\Desktop\suggestions_summary.xlsx', sheet_name='suggestions')
df_output = pd.DataFrame(suggestion_sum_list)
df_output.to_excel(r'C:\Users\user\Desktop\suggestions_list.xlsx', sheet_name='test', index=0, header=0)
print('花費時間：%.2f秒' % (time.time()-t0))


