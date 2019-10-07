# 行政院環境保護署 空氣品質監測網
# https://taqm.epa.gov.tw/taqm/tw/YearlyDataDownload.aspx
from pathlib import Path  #python內建檔案處理的模組
import matplotlib   #繪圖套件
import matplotlib.pyplot as plt
import pandas as pd #數據分析處理套件
import datetime

print('初始化程式....')
IAQI = {'O3': [(0, 108), (109, 140), (141, 170),
               (171, 210), (211, 400)],
        'PM2.5': [(0, 12), (13, 35.4), (35.5, 55.4), (55.5, 150.4), (150.5, 250.4), (250.5, 350.4), (350.5, 500.5)],
        'PM10': [(0, 54), (55, 154), (155, 254), (255, 354), (355, 424), (425, 504), (505, 600)],
        'CO': [(0, 5.038), (5.039, 10.763), (10.764, 14.198), (14.199, 17.633), (17.644, 34.35), (34.351, 46.258), (46.259, 57.708)],
        'SO2': [(0, 91.7), (91.8, 196.5), (196.6, 484.7), (484.8, 799.1), (799.2, 1582.5), (1582.6, 2106.5), (2106.6, 2630.5)],
        'NO2': [(0, 99.64), (99.65, 188), (188.01, 676.8), (676.81, 1220), (1220.01, 2350), (2350.01, 3100), (3100.01, 3850)]
        }


# 去除奇怪的符號
def remove(v):
    if isinstance(v, str):
        if '#' in v:
            return v[:-1]
        if 'x' in v:
            return v[:-1]
        if '*' in v:
            return v[:-1]
        if 'A' in v:
            return v[:-1]
    return v


# 計算IAQIP
def iaqip(row):
    CP = row[3:].astype(float).mean()
    for IAQILo, IAQIHi in IAQI[row['測項']]:
        if IAQILo >= CP and CP <= IAQIHi:
            BPHi = row[3:].astype(float).max()
            BPLo = row[3:].astype(float).min()
            IAQIP = (IAQIHi - IAQILo)/(BPHi-BPLo)*(CP - BPLo) + IAQILo
            return IAQIP

print('檔案讀取中....')
if Path('AQI.feather').exists():
    df = pd.read_feather('AQI.feather', nthreads=4)
else:
    df = pd.DataFrame()
    path = list(Path('.').glob('**/*.xls'))
    for f in path:
        data = pd.read_excel(f)
        df = df.append(data)

    print('資料清理中....')
    df = df[(df['測項'] == "PM2.5") | (df['測項'] == "PM10")
            | (df['測項'] == "SO2") | (df['測項'] == "NO2")
            | (df['測項'] == "O3") | (df['測項'] == "CO")]
    df = df.applymap(remove)

    print('計算AQI...')
    df['AQI'] = df.apply(iaqip, axis=1)

print('繪圖中...')
df2 = df.groupby(['日期']).max(numeric_only=True).reset_index()
df2['日期'] = pd.to_datetime(df2['日期'])
# df2.plot.line(x='日期',y=['AQI'])
# print(df2[df2['日期'].dt.year == 2018] )
# df2[df2['日期'].dt.year == 2018].plot.line(x=('1','2','3','4','5','6','7','8','9','10','11','10'), y=['AQI'])
# df2[df2['日期'].dt.year == 2016].plot.line(y=['AQI'])
# df2[df2['日期'].dt.year == 2015].plot.line(y=['AQI'])
df2[df2['日期'].dt.year == 2017].plot.line(x='日期',y=['AQI'])
plt.legend(prop=matplotlib.font_manager.FontProperties(fname='C:\Windows\Fonts\msjh.ttc'))
print('完成!!')
df2.to_feather('AQI.feather')
plt.show()



# df['24小時平均'] = df.iloc[:, 3:].astype(float).mean(1)
# df['最大值'] = df.max(1)
# df['最小值'] = df.min(1)
# df.groupby(['日期']).max(numeric_only=True)['IAQIP']
# plot.line(x='日期',y='IAQIP')
# df.apply(test, axis=1)
# print(df.iloc[:3])


# print()
# print(df.iloc[0:3, 3:].max(1))
# iaqi.append([])

# print()


# print(df.shape[0]) #有幾筆資料
# IAQI[df['']]
# print()


# print(df[df['日期'] == '2018/01/01'])
# print(df.iloc[0:].str.isnumeric() )
# print(df.iloc[13:20,3:])
# print(int('56#'))
