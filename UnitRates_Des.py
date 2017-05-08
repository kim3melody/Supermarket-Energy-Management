# -*- coding: utf-8 -*-
"""
Created on Thu Sep 15 09:46:58 2016

@author: czhang0914
"""

import pandas as pd
import os

folder = r'C:\Users\czhang0914\Downloads\Optima Data'
os.chdir(folder)

Invoice=pd.read_excel('Jul2015-Jun2016 Original Data.xlsx',header=3)
Meters=pd.read_excel('Unit Rates Items.xlsx',sheetname='Sheet1',header=None)

Meters=Meters.iloc[:,0].tolist()
dic={}

for i,n in enumerate(Meters):
    dic[i]=Invoice[n]

dic[0]=pd.to_datetime(dic[0],format='%b %Y')

ua=pd.concat(dic,axis=1)
de=ua.groupby(['Bill Period']).sum()
de_avg=ua.groupby(['Bill Period']).mean()
de_count=ua.groupby(['Bill Period']).count()

Normal_Charge=de.iloc[:,1:8].sum(axis=1)
Additional_Charge=de.iloc[:,8:39].sum(axis=1)
Normal_UnitRate=Normal_Charge/de.iloc[:,0]
Raise_UnitRate=Additional_Charge/de.iloc[:,0]

Summary=pd.concat([Normal_Charge,Additional_Charge,Normal_UnitRate,Raise_UnitRate],axis=1)
Summary.columns=(['Normal_Charge','Additional_Charge','Normal_UnitRate','Raise_UnitRate'])
Summary['All_Including_UR']=Summary['Normal_UnitRate']+Summary['Raise_UnitRate']
Summary['Charge_Percentage Increase']=Summary['Additional_Charge']/Summary['Normal_Charge']

with pd.ExcelWriter('UnitRate_Des.xlsx') as writer:
    Summary.to_excel(writer,sheet_name='Summary')
    de.T.to_excel(writer,sheet_name='Sum')
    de_avg.T.to_excel(writer,sheet_name='Avg')
    de_count.T.to_excel(writer,sheet_name='Data Point Count')