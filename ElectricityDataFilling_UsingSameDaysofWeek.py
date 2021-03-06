# -*- coding: utf-8 -*-
"""
Created on Wed Oct  5 14:04:01 2016

@author: czhang0914
"""

import pandas as pd
import os

path=r'C:\Users\czhang0914\Desktop\Monthly Report Data\Apr data 2017'
ElecData=r'M105 Apr data 2017.xlsx'
TargetFileName=r'ElectricityDataEstimation M105 Apr2017.xlsx'

os.chdir(path)
path = os.getcwd()

Elec=pd.read_excel(ElecData,sheetname='Elec')

Elec=Elec[['OPTIMA Half Hourly DATA', 'READING DATE','Daily Total']]
Elec=Elec[Elec['OPTIMA Half Hourly DATA'].str.isdigit()]
Pivoted=Elec.pivot(index='READING DATE',columns='OPTIMA Half Hourly DATA',values='Daily Total')
Pivoted['DayofWeek']=Pivoted.index.weekday

#NaFilled=Pivoted.groupby('DayofWeek').transform(lambda x: x.fillna(x.mean()))

NaFilled=Pivoted.groupby('DayofWeek').fillna(method="ffill",axis=0)

with pd.ExcelWriter(TargetFileName) as writer:
    Pivoted.to_excel(writer,sheet_name='Pivoted') 
    NaFilled.to_excel(writer,sheet_name='NaFilled') 
   