# -*- coding: utf-8 -*-
"""
Created on Wed Oct  5 14:04:01 2016

@author: czhang0914
"""

import pandas as pd
import os

path=r'C:\Users\czhang0914\Desktop\Monthly Report Data\Dec data 2016'
GasData=r'M395 Dec data 2016.xlsx'
TargetFileName=r'GasDataEstimation M395 Dec2016.xlsx'

os.chdir(path)
path = os.getcwd()

Gas=pd.read_excel(GasData,sheetname='Gas')

Gas=Gas[['OPTIMA Half Hourly DATA', 'READING DATE','Daily Total']]
Pivoted=Gas.pivot(index='READING DATE',columns='OPTIMA Half Hourly DATA',values='Daily Total')
Pivoted['DayofWeek']=Pivoted.index.weekday

#NaFilled=Pivoted.groupby('DayofWeek').transform(lambda x: x.fillna(x.mean()))

NaFilled=Pivoted.groupby('DayofWeek').fillna(method="ffill",axis=0)

with pd.ExcelWriter(TargetFileName) as writer:
    Pivoted.to_excel(writer,sheet_name='Pivoted') 
    NaFilled.to_excel(writer,sheet_name='NaFilled') 
   