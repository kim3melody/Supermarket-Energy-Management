# -*- coding: utf-8 -*-
"""
Created on Fri Feb  3 10:01:14 2017

@author: czhang0914
"""

# -*- coding: utf-8 -*-
"""
Created on Tue Mar 08 09:44:44 2016

@author: czhang0914
"""

import os
import pandas as pd


# Set the location of CDD files, you'll find the result in the same folder
CDDfolder=r'C:\Users\czhang0914\Desktop\Monthly Report Data\Feb data 2016\DD files Feb 2016'
start='20170101'
end='20170131'
TargetFileName=r'2016Feb.xlsx'


# Date generation
startd = start[:4]+'/'+start[4:6]+'/'+start[6:8]
endd = end[:4]+'/'+end[4:6]+'/'+end[6:8]
# Change working directory to collect CDD values from each file
os.chdir(CDDfolder)
path = os.getcwd()
files = os.listdir(path)

cddfiles = [f for f in files if "CDD" in f]
df = pd.DataFrame(index = pd.date_range(start = startd, end = endd,freq='D'))

for i in cddfiles:
    ws=i[:-19]
    data = pd.read_csv(i,header=6,usecols=['Date','3'],index_col=['Date'])
#    data.index = pd.to_datetime(data.index,format = '%Y-%m-%d')
    data.index = pd.to_datetime(data.index,format = '%m/%d/%Y')
    data=data.rename(columns = {'3':ws})
    df=pd.concat([df,data],axis=1)


df=df.loc[start:end]
s=df.append(df.sum(numeric_only=True), ignore_index=True)
s=s.loc[[7]]
s=s.T
s.columns=[start]
s.index.names = ['2016 CDD']

with pd.ExcelWriter(TargetFileName) as writer:
    s.to_excel(writer,sheet_name='weekly cdd total') 
    df.to_excel(writer,sheet_name='check CDD items') 
   