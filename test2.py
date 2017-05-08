# -*- coding: utf-8 -*-
"""
Created on Tue Mar 08 09:44:44 2016

@author: czhang0914
"""

import os
import pandas as pd


os.chdir('C:\Users\czhang0914\Desktop\Feb data 2016\DD files Feb 2016')
path = os.getcwd()
files = os.listdir(path)

cddfiles = [f for f in files if "CDD" in f]
