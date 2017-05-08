# -*- coding: utf-8 -*-
"""
Created on Wed Apr 26 16:32:14 2017

@author: czhang0914
"""

import pandas as pd
import os

path = r'C:\Users\czhang0914\Desktop'
file = r'excel.xlsx'

os.chdir(path)
res ={}

for i in range(24):
    tmp = pd.read_excel(file, sheetnaem = str(i), header = [2,3,4] )
    tmp.index = tmp.index.map(lambda x: x[-8:])
    res[i] = tmp
    
    