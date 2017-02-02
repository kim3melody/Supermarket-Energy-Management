# -*- coding: utf-8 -*-
"""
Created on Tue Mar 01 14:05:10 2016

@author: czhang0914
"""

# Change point variable calculation for F 64 and 68

import pandas as pd

TargetFileName='CHPTcalc Dec 2012.xlsx'

meanTemp=pd.read_excel('Web_Query_DownloadWeather_Custom.xlsm',sheetname='Summary')
meanTemp=meanTemp.dropna(axis=0,how='any')
CHPT64=meanTemp.apply(lambda x: 64-x)
CHPT64[CHPT64<0]=0
CHPT68=meanTemp.apply(lambda x: 68-x)
CHPT68[CHPT68<0]=0

# datetime index to group weekdays and sundays (0-6 represents Mon-Sun)
di=pd.DatetimeIndex(meanTemp.index)
CHPT64.insert(0,'dayofweek',di.weekday)
CHPT64=CHPT64.set_index('dayofweek')
CHPT68.insert(0,'dayofweek',di.weekday)
CHPT68=CHPT68.set_index('dayofweek')


weekdaysum64=CHPT64[CHPT64.index.isin(range(6))].sum()
weekdaysum64.name='weekdays sum CHPT 64'
sundaysum64=CHPT64[CHPT64.index.isin([6])].sum()
sundaysum64.name='sundays sum CHPT 64'

weekdaysum68=CHPT68[CHPT68.index.isin(range(6))].sum()
weekdaysum68.name='weekdays sum CHPT 68'
sundaysum68=CHPT68[CHPT68.index.isin([6])].sum()
sundaysum68.name='sundays sum CHPT 68'

Summary=pd.concat([weekdaysum64,sundaysum64,weekdaysum68,sundaysum68],axis=1)


with pd.ExcelWriter(TargetFileName) as writer:
    Summary.to_excel(writer,sheet_name='CHPT') 
    meanTemp.to_excel(writer,sheet_name='MeanTempRecords') 
   