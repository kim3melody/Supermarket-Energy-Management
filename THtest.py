# -*- coding: utf-8 -*-
"""
Created on Wed Dec  7 11:29:16 2016

@author: czhang0914
"""

import pandas as pd

DailyTH = pd.read_excel('Trading Hours calculated Sheet - 6th Nov 2016. Remake V4.xlsx',sheetname='TradingHours-Daily',index_col=[0,1,2,3])