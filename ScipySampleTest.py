# -*- coding: utf-8 -*-
"""
Created on Tue Sep 27 15:33:24 2016

@author: czhang0914
"""

import pandas as pd

data=pd.read_excel('brain_size.xlsx')

groupby_gender = data.groupby('Gender')
