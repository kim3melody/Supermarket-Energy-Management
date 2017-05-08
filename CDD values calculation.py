# -*- coding: utf-8 -*-
"""
Created on Tue Mar 08 09:44:44 2016

@author: czhang0914
"""

import os
import pandas as pd

CDDfolder=r'C:\Users\czhang0914\Desktop\Monthly Report Data\Nov data 2016\DD files Nov 2016'

os.chdir(CDDfolder)
path = os.getcwd()
files = os.listdir(path)

cddfiles = [f for f in files if "CDD" in f]
tpl=pd.read_csv(cddfiles[0],header=5,usecols=['Date','CDD'])
index=pd.DatetimeIndex(tpl.Date)
df=pd.DataFrame(index=index)

for i in cddfiles:
    ws=i[:-13]
    data=pd.read_csv(i,header=5,usecols=['Date','CDD'])
    data=data.set_index(pd.DatetimeIndex(data.Date))
    data=data.drop('Date',axis=1)
    data=data.rename(columns = {'CDD':ws})
    df=pd.concat([df,data],axis=1)
    
df.to_excel('NovDD.xlsx')