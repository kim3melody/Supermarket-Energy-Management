# -*- coding: utf-8 -*-
"""
Created on Fri Apr 21 15:18:34 2017

@author: czhang0914
"""

import os
import zipfile
import pandas as pd

DataLog = {}
Events = {}
zipped = []
DataLog = {}
Events = {}

rootdir = r'C:\Users\czhang0914\Desktop\Logoplaste Netherlands\Blowing\Husky Data'
os.chdir(rootdir)

# one time extract all process

#for subdir, dirs, files in os.walk(rootdir):
#    for file in files:
#        path = os.path.join(subdir, file)
#        zipped.append(path)
#        if '.Zip' in path:
#            tmp = zipfile.ZipFile(path)
#            tmp.extractall(os.path.join(subdir, 'unzipped'))
#        

for subdir, dirs, files in os.walk(rootdir):
    for file in files:
        path = os.path.join(subdir, file)
        if 'DataLog.csv' in path:
            print('reading...' + path)
            DataLog[subdir] = pd.read_csv(path, header = [0, 1], tupleize_cols = True, encoding = 'cp1252', parse_dates =[0])
          
        elif 'Events.csv' in path:
            print('reading...' + path)
            tmp = pd.read_csv(path, names = ['Date/Time', 'Type', 'Source', 'Description', 'ExtraDes'], header= None, encoding = 'cp1252', parse_dates =[0])
            tmp.drop(0, axis =0, inplace =True)
            Events[subdir] = tmp

for key, item in DataLog.items():
    key = key[51:]

for key, item in Events.items():
    key = key[51:]

os.chdir(r'C:\Users\czhang0914\Desktop\Logoplaste Netherlands\Organized\DataLog')
for key, item in DataLog.items():
    print('writing...'+key)
    key = key[51:]
    key = key.replace('\\','-')
    item.to_excel(key+'.xlsx')
    
os.chdir(r'C:\Users\czhang0914\Desktop\Logoplaste Netherlands\Organized\Events')
for key, item in Events.items():
    print('writing...'+key)
    key = key[51:]
    key = key.replace('\\','-')
    item.to_excel(key+'.xlsx')