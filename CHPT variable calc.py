# -*- coding: utf-8 -*-
"""
Created on Tue Mar 01 14:05:10 2016

@author: czhang0914
"""

# Change point variable calculation for F 64 and 68

import pandas as pd
import os

path=r'C:\Users\czhang0914\desktop'
TargetFileName='CHPTcalc Apr 2013.xlsx'
os.chdir(path)
#meanTemp=pd.read_excel('Jan_2017.xlsx',sheername = 'Jan_2017')
meanTemp=pd.read_csv('apr_2013.csv',index_col = 0)


#meanTemp=meanTemp.dropna(axis=0,how='any')      # drop null data, check and fix original data file before this step
CHPT64=meanTemp.apply(lambda x: 64-x)           #  map function to apply 64-x for items in the list
CHPT64[CHPT64<0]=0                              # discard negative result
CHPT68=meanTemp.apply(lambda x: 68-x)
CHPT68[CHPT68<0]=0

# datetime index to group weekdays and sundays (0-6 represents Mon-Sun)

di=pd.DatetimeIndex(meanTemp.index)                 # set the index using date
CHPT64.insert(0,'dayofweek',di.weekday)             # insert a coulumn to indicate what day  it is in a week
CHPT64=CHPT64.set_index('dayofweek')
CHPT68.insert(0,'dayofweek',di.weekday)
CHPT68=CHPT68.set_index('dayofweek')


weekdaysum64=CHPT64[CHPT64.index.isin(range(6))].sum()  # 0-5   Mon- Sat
weekdaysum64.name='weekdays sum CHPT 64'
sundaysum64=CHPT64[CHPT64.index.isin([6])].sum()  # 6  Sun
sundaysum64.name='sundays sum CHPT 64'

weekdaysum68=CHPT68[CHPT68.index.isin(range(6))].sum()
weekdaysum68.name='weekdays sum CHPT 68'
sundaysum68=CHPT68[CHPT68.index.isin([6])].sum()
sundaysum68.name='sundays sum CHPT 68'

Summary=pd.concat([weekdaysum64,sundaysum64,weekdaysum68,sundaysum68],axis=1)


with pd.ExcelWriter(TargetFileName) as writer:
    Summary.to_excel(writer,sheet_name='CHPT') 
    meanTemp.to_excel(writer,sheet_name='MeanTempRecords') 
   