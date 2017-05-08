# -*- coding: utf-8 -*-
"""
Created on Wed Aug 17 09:49:44 2016

@author: czhang0914
"""

import pandas as pd
import os

CHPfolder=r'C:\Users\czhang0914\Desktop\Weekly Reports to Dan Mitchill\Weekly Exec\week 33 BM reports\wk33 CHP'
sw='20160912'   # start of the week
ew='20160918'   # end of the week
TargetFileName=r'CHPwk32.xlsx'

# Change working directory to collect CHP values from each file
os.chdir(CHPfolder)
#path = os.getcwd()
files = os.listdir(CHPfolder)

daily = {}

for i in files:
    ws=i[13:-5]
    ws=pd.to_datetime(ws,format='%d-%m-%y')
    data=pd.read_excel(i,sheetname='Overview',header=0)
    data=data[pd.notnull(data.index)]
    data=data.iloc[:,6]
    daily[ws]=data
    #data=data.rename(columns={'Daily kW(e)':ws})
    

summary=pd.concat(daily,axis=1)
weekly=summary.loc[:,sw:ew]
weekly['Total']=weekly.sum(axis=1)

weekly.to_excel(TargetFileName)