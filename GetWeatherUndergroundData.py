# -*- coding: utf-8 -*-
"""
Created on Mon Apr 04 16:20:43 2016

@author: czhang0914
"""


import pandas as pd
import io
import urllib


# progress bar
import sys

def ProgressBar(val,end_val, bar_length=20):
    percent = float(val) / end_val
    hashes = '#' * int(round(percent * bar_length))
    spaces = ' ' * (bar_length - len(hashes))
    sys.stdout.write("\rProgress: [{0}] {1}%".format(hashes + spaces, int(round(percent * 100))))
    sys.stdout.flush()
    

# Date format "yyyy/mm/dd"
sdate='2016/12/1'
edate='2016/12/31'
sdate=sdate.split('/')
edate=edate.split('/')


wslist=pd.read_excel('WeatherStationList.xlsx')
count=0
list_len=len(wslist)
wslist=list(wslist.iloc[:,0])
WeatherData={}

for i in wslist:    
    # parse destination url
    #time.sleep(10)
    base_url='https://www.wunderground.com/history/airport/{WeatherStation}/{syear}/{smonth}/{sday}/CustomHistory.html?dayend={eday}&monthend={emonth}&yearend={eyear}&req_city=&req_state=&req_statename=&reqdb.zip=&reqdb.magic=&reqdb.wmo=&format=1'
    url=base_url.format(WeatherStation=i,syear=sdate[0],smonth=sdate[1],sday=sdate[2],eyear=edate[0],emonth=edate[1],eday=edate[2])

    s=urllib.request.urlopen(url,data=None)
    WeatherData[i]=pd.read_csv(io.StringIO(s.read().decode('utf-8')),index_col=0,usecols=[0,1,2,3])
    count+=1
    ProgressBar(count,list_len)
   
wd=pd.concat(WeatherData,axis=1) 
#wd.to_csv('nov 20161215.csv')