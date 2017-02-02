# -*- coding: utf-8 -*-
"""
Created on Tue Mar 08 09:44:44 2016

@author: czhang0914
"""

import os
import pandas as pd


# Set the location of CDD files, you'll find the result in the same folder
CDDfolder=r'C:\Users\czhang0914\Downloads\CDD'
startofweek=20170123
endofweek=20170129
TargetFileName=r'Week52cddSUM.xlsx'

# Change working directory to collect CDD values from each file
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

startofweek=str(startofweek)
endofweek=str(endofweek)
df=df.loc[startofweek:endofweek]
s=df.append(df.sum(numeric_only=True), ignore_index=True)
s=s.loc[[7]]
s=s.T
s.columns=[startofweek]
s.index.names = ['2016 CDD']

with pd.ExcelWriter(TargetFileName) as writer:
    s.to_excel(writer,sheet_name='weekly cdd total') 
    df.to_excel(writer,sheet_name='check CDD items') 
   