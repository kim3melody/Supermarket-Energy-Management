# -*- coding: utf-8 -*-
"""
Created on Fri Dec 16 09:38:17 2016

@author: czhang0914
"""


import pandas as pd
import os

path=r'C:\Users\czhang0914\Desktop\Gibraltor Barracks\Gibraltor Barracks'
os.chdir(path)
files=os.listdir(path)

res={}
# loop through each file
for f in files:
    f_name = f[0:-5]
    df=pd.read_excel(f,header=None)
    table_names = ['Address / Network / IO','Sensor List','Digital Inputs List','Knobs List','Switch List']
    groups = df[14].isin(table_names).cumsum()
    tables = {g.iloc[0,14]: g.iloc[2:] for k,g in df.groupby(groups) if k>0}
              
    # rename dictionary keys
    tables['Driver'] =tables.pop('Switch List')
    tables['Driver'] = tables['Driver'][:-2]
    tables['Switch'] =tables.pop('Knobs List')
    tables['Knob'] =tables.pop('Digital Inputs List')
    tables['Dig In'] =tables.pop('Sensor List')
    tables['Sensor List'] =tables.pop('Address / Network / IO')
    
    f_summary={} 
    # reset headers for each table
    for i in tables:
        f_summary[i] = pd.DataFrame([0,0],index=['Total Points Available','Total Points After Filtered'])
        try:
            new_header = tables[i].iloc[0]
            tables[i] = tables[i][1:]
            tables[i].columns = new_header
        except:
            pass
            
        # count the 'Label' column
        try:
            items = tables[i]['Label'].fillna('spare')
            f_summary[i].loc['Total Points Available'] = len(items)
            f_summary[i].loc['Total Points After Filtered'] = len([x for x in items if 'spare' not in x.lower()])
        except:          
            print(f+' '+i+' Error: no column named as Label')
    
    site_summary=pd.concat(f_summary,axis=1)
    site_summary.columns=site_summary.columns.droplevel(level=1)
    site_summary=site_summary[['Sensor List','Dig In','Knob','Switch','Driver']]
    res[f]=site_summary
    
result=pd.concat(res)
result.to_excel('BMS points count 20161216.xlsx')