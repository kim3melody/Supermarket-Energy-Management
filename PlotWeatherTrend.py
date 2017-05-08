# -*- coding: utf-8 -*-
"""
Created on Thu Feb 25 14:43:21 2016

@author: czhang0914
"""

import pandas as pd
import matplotlib.pyplot as plt

xl=pd.ExcelFile('Web_Query_DownloadWeather_Custom.xlsm')
tabs=xl.sheet_names
date=xl.parse(tabs[2]).ix[:,0]

data=pd.DataFrame(date)
for i in tabs[2:]:
    data[i]=xl.parse(i).ix[:,2]
    
#data=data.dropna(how='any')
data=data.set_index(data.ix[:,0])
data=data.drop(data.columns[0],axis=1)
summary=data.T.describe()
P=summary.loc['top']
plt.figure()
P.plot(title='CDD trend',legend=True)



