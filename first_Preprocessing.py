import pandas as pd
import numpy as np
from scipy import spatial
import warnings
warnings.filterwarnings('ignore')

sheet_name=['中电投望洋台东风电一场','特中老君庙风电一场','龙源望洋台西风电一场',
           '天润十三间房风电一场','华电小草湖北风电一场','中电投玛依塔斯风电二场',
           '中广核达坂城风电场','粤水电布尔津风电一场','国电阿拉山口风电场',
           '大唐若羌风电一场']
data=pd.read_excel('e:/2018风电数据.xlsx',sheet_name=sheet_name[9])

col=['测风塔10m风速(m/s)',
 '测风塔30m风速(m/s)',
 '测风塔50m风速(m/s)',
 '测风塔70m风速(m/s)',
 '轮毂高度风速(m/s)',
 '测风塔10m风向(°)',
 '测风塔30m风向(°)',
 '测风塔50m风向(°)',
 '测风塔70m风向(°)',
 '轮毂高度风向(°)',
 '温度(°)',
 '气压(hPa)',
 '实际发电功率（mw）']

cols=list(np.arange(3,16))

for i,f in enumerate(col):
    isx=data[data[f]==0].index[0]
    c=cols.copy()
    del c[i]
    vec1=data.iloc[isx,c].values.astype(float)
    cos_distance=[]
    for j in np.arange(data.shape[0]):
        vec2=data.iloc[j,c].values.astype(float)
        cos_distance.append(1-spatial.distance.cosine(vec1,vec2))
    cos_d=pd.DataFrame(cos_distance)
    index_replace=cos_d.sort_values(by=0,ascending=False).index[50]
    values_r=data.iloc[index_replace,cols[i]]
    while values_r ==0:
        index_replace+=1
        values_r=data.iloc[index_replace,cols[i]]
    data[f].replace([0],[values_r],inplace=True)

n=0
for i in np.arange(data.shape[0]):
    a=data.iloc[i,3]
    b=data.iloc[i,4]
    c=data.iloc[i,5]
    d=data.iloc[i,6]
    if a<b<c<d:
        n+=1
        continue
    else:
        data.iloc[i,3]=data.iloc[i-1,3]
        data.iloc[i,4]=data.iloc[i-1,4]
        data.iloc[i,5]=data.iloc[i-1,5]
        data.iloc[i,6]=data.iloc[i-1,6]
print('替换行数：',data.shape[0]-n)

m=0
for i in np.arange(data.shape[0]):
    f=abs(data.iloc[i,8:12].diff().values[1:].astype(float))>30
    if f.sum()==0:
        m+=1
        continue
       
    else:
        data.iloc[i,8]=data.iloc[i-1,8]
        data.iloc[i,9]=data.iloc[i-1,9]
        data.iloc[i,10]=data.iloc[i-1,10]
        data.iloc[i,11]=data.iloc[i-1,11]
print('替换行数：',data.shape[0]-m)
xlsx=pd.ExcelWriter('e:/2018风电数据--新.xlsx')
data.to_excel(xlsx, sheet_name=sheet_name[0])
xlsx.close()