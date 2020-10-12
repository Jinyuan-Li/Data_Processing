import re
import json

with open(r'C:\Users\user\Desktop\html\public_html\javascripts\TOWN_MOI.json',encoding='utf8') as f:
    data = json.load(f)

sTOWN_A,sTOWN_B,sTOWN_C,sTOWN_D,sTOWN_E,sTOWN_F,sTOWN_G,sTOWN_H,sTOWN_I,sTOWN_J,sTOWN_K,sTOWN_M,sTOWN_N,sTOWN_O,sTOWN_P,sTOWN_Q,sTOWN_T,sTOWN_U,sTOWN_V,sTOWN_W,sTOWN_X,sTOWN_Z = '','','','','','','','','','','','','','','','','','','','','',''
TOWN_list = [sTOWN_A,sTOWN_B,sTOWN_C,sTOWN_D,sTOWN_E,sTOWN_F,sTOWN_G,sTOWN_H,sTOWN_I,sTOWN_J,sTOWN_K,sTOWN_M,sTOWN_N,sTOWN_O,sTOWN_P,sTOWN_Q,sTOWN_T,sTOWN_U,sTOWN_V,sTOWN_W,sTOWN_X,sTOWN_Z]
COUNTYID_list = ['A','B','C','D','E','F','G','H','I','J','K','M','N','O','P','Q','T','U','V','W','X','Z']

for i in range(len(data['features'])):
    for j, k in enumerate(COUNTYID_list):
        if data['features'][i]['properties']['COUNTYID'] == k:
            aaa = re.findall(r"'value': (.+?), 'COUNTYID'",str(data['features'][i]))
            bbb = 'data[year_index].TOWN_'+str(data['features'][i]['properties']['TOWNID'])
            ccc = str(data['features'][i]).replace(aaa[0], bbb)
            TOWN_list[j] += ccc + ',\n'

zzz = ''
for xxx in TOWN_list:
    yyy = 'statesData = {"type":"FeatureCollection", "features": [' + xxx + ']}\n\n\n\n\n\n'
    zzz += yyy

f = open(r'C:\Users\user\Desktop\test.txt', 'w', encoding='utf8')
f.write(zzz)
f.close()



