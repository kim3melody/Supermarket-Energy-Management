# -*- coding: utf-8 -*-
"""
Created on Tue Nov  8 11:45:13 2016

@author: czhang0914
"""

import os
import pandas as pd

CDDfolder=r'C:\Users\czhang0914\Desktop\Monthly Report Data\Feb data 2016\DD files Feb 2016'

os.chdir(CDDfolder)
path = os.getcwd()
files = os.listdir(path)

cddfiles = [f for f in files if "CDD" in f]
tpl=pd.read_csv(cddfiles[0],header=6,usecols=['Date','3'])
index=pd.DatetimeIndex(tpl.Date)
df=pd.DataFrame(index=index)

for i in cddfiles:
    ws=i[:-19]
    data=pd.read_csv(i,header=6,usecols=['Date','3'])
    data=data.set_index(pd.DatetimeIndex(data.Date))
    data=data.drop('Date',axis=1)
    data=data.rename(columns = {'3':ws})
    df=pd.concat([df,data],axis=0)
    
df.to_excel('DDintegration.xlsx')