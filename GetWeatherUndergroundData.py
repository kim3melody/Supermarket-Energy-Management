# -*- coding: utf-8 -*-
"""
Created on Mon Apr 04 16:20:43 2016

@author: czhang0914
"""


import pandas as pd
import io
import urllib
import os

path=r'C:\WinPython-64bit-3.4.4.5Qt5\zcc'
output_name=r'Jan_2017.csv'

sdate='2017/1/1'
edate='2017/1/31'


os.chdir(path)
# progress bar
import sys

def ProgressBar(val,end_val, bar_length=20):
    percent = float(val) / end_val
    hashes = '#' * int(round(percent * bar_length))
    spaces = ' ' * (bar_length - len(hashes))
    sys.stdout.write("\rProgress: [{0}] {1}%".format(hashes + spaces, int(round(percent * 100))))
    sys.stdout.flush()
    
# Date format "yyyy/mm/dd"
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
    WeatherData[i]=pd.read_csv(io.StringIO(s.read().decode('utf-8')),index_col=0,usecols=[0,2])
    WeatherData[i]['Mean TemperatureC']=WeatherData[i]['Mean TemperatureC']*1.8+32
    WeatherData[i] = WeatherData[i].rename(columns={'Mean TemperatureC':'Mean TemperatureF'})
    WeatherData[i].index=pd.to_datetime(WeatherData[i].index,format='%Y-%m-%d')
    count+=1
    ProgressBar(count,list_len)
   
wd=pd.concat(WeatherData,axis=1) 
wd.to_csv(output_name)
