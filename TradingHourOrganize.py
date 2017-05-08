# -*- coding: utf-8 -*-
"""
Created on Thu Mar 24 15:27:42 2016

@author: czhang0914
"""

import pandas as pd

th=pd.read_csv('TH.csv',index_col=0,parse_dates=[0])
weeklyth=th.resample('W',how='sum').reset_index()