# -*- coding: utf-8 -*-
"""
Created on Wed Jun 22 14:22:17 2016

@author: czhang0914
"""

import pandas as pd
import sys  
reload(sys)  
sys.setdefaultencoding('utf8')

FoodLoad=pd.read_csv('Food Load Data 20160401-20160430.csv',header=None)
Sorted=FoodLoad.sort_values(by=[0,2,4],ascending=[1,1,1])
First300=Sorted.groupby([0,2]).head(300)
Summary=First300.groupby([0,1,2]).mean()

Summary.to_excel('xxx.xlsx')