#!/usr/bin/python2.7
# -*- coding: utf-8 -*-
"""
Created on Mon Jul 13 13:59:56 2015

@author: jzhu0922
"""

import numpy 
import matplotlib.pyplot as plt

import pandas as pd 
from pandas import DataFrame 
from pandas import Series
#from datetime import datetime

def is_open_at(x, op, ed, op_S, ed_S, delay=0.0):
    t = x.hour + x.minute / 60.0
    if x.weekday() == 6:
        return (op_S-delay <= t) and (t < ed_S + delay)
    else:
        return (op-delay <= t) and (t < ed + delay)

def is_lighting_ecm(x, name, thld, mult=1, add=0):
    return (x.open == False) and (x[name] > mult * thld + add)
    
def is_lighting_ecm2(x, name, thld, ave, mult=0):
    return (x.open == False) and (x[name] > mult * ave + (1 - mult) * thld )

def event_to_ecm(df, site, thld, limit = 0):
    ecms = {'start':[], 'end':[], 'saving':[]}
    i = 0
    while i < len(df):
        if df.ecm[i]:
            j = i
            save = 0
            while df.ecm[i]:
                i = i + 1
                save = save + (df[site][i] - thld) * 5 / 60
            if (i-j >= limit) and (df[site][j] != df[site][i-1] or df[site][j] != df[site][j+1]):
                ecms['start'].append(df.index[j])
                ecms['end'].append(df.index[i])
                ecms['saving'].append(save)
        else:
            i = i + 1
    return ecms

def plot_ecm(data, start_str, end_str, site, result, thld, daytime):
    p = data.ix[start_str:end_str, site].plot(title = site)
    p.set_ylabel('Lighting Demand Power 5 min (kw)')
    p.set_xlabel('Timestamp')
    p.axhline(y = thld, color='r', linewidth = 2)
    p.axhline(y = daytime, color='r', linewidth = 2)
    p.axhline(y = 0.6*thld + 0.4*daytime, color='g', linestyle='--', linewidth = 2)
    if site in result:
        ecms = result[site]
        for i in range(len(ecms['start'])):
            p.axvspan(ecms['start'][i], ecms['end'][i], facecolor='y', alpha=0.4)

def output_ecm(result):
    result_list = []
    for k,v in result.iteritems():
        df_ecm = DataFrame(v)
        df_ecm['site'] = k
        result_list.append(df_ecm)
    df= pd.concat(result_list)
    df = df[['site', 'start', 'end', 'saving']]
    df.set_index('site', inplace=True)
    return df

def simulation(data,df_info):  
    #initialize result data structure
    result = {}
    site_count = 0
    #loop through sites
    for site in data.columns:
        s = df_info.ix[site]    #get info, better if we can have siteID 
        df = DataFrame(data[site])
        # t is a series of time, check the open time and alert
        t = Series(df.index).apply(is_open_at, op=s.OpenWeekday, ed=s.CloseWeekday, op_S=s.OpenWeekend, ed_S=s.CloseWeekend, delay=0.5)
        t.index = df.index
        
        df['open'] = t
#        df['ecm'] = df.apply(is_lighting_ecm, name=site, thld=s.Threshold, mult=1.3, add = 3, axis=1)
        df['ecm'] = df.apply(is_lighting_ecm2, name=site, thld=s.Threshold, ave=s.DaytimeAve, mult=0.4, axis=1)
        df.ix[-1,'ecm'] = False   #make sure the alarm will stop, i know i am lazy
           
        if df['ecm'].any():
            item = event_to_ecm(df,site,s.Threshold,limit=6)
            if len(item['start'])>0: result[site] = item
    
        site_count = site_count + 1
        print str(site_count) + ' of ' + str(len(data.columns)) + ' : ' + site
    
    return result


if __name__ == "__main__":
    #read the site info=================================================
    df_info = pd.read_csv('./file/site_info.csv',sep=',')
    #df_info['SiteID'] = df_info['SiteID'].apply(str)
    df_info.set_index('SiteName', inplace=True)    #maybe just use sitename as index, as long it is verified
    
    #read the demand power data=========================================
    data = pd.read_csv('./file/SepPivot.csv', sep=',')
    data['time'] = pd.to_datetime(data['time'], dayfirst=False)
    data.set_index('time', inplace=True)
    data.replace(to_replace=0, value=numpy.NaN, inplace=True) #some 0s seems like missing value
    data = data.resample('5min',fill_method='ffill')
    data.fillna(method='ffill', inplace=True) 
    
    result = simulation(data, df_info)  #get the result from simulation
        
    r1 = output_ecm(result)   
    r2 = r1.merge(df_info, left_index=True, right_index=True)
    r2.to_csv('./Sept.csv')    
    
#==============================================================================
#plot ecm sites only
#    sites = r2.index.unique()
#    plt.rc('figure', figsize=(40, 25)) #need %pylab
#    for i in range(len(sites)):
#        plot_ecm(data,'20150601', '20150624', sites[i], result, df_info.ix[sites[i]].Threshold, df_info.ix[sites[i]].DaytimeAve)
#        plt.savefig("./fig/"+sites[i]+".jpeg", dpi = 150)
#        plt.close()

#==============================================================================
#plot all the sites
plt.rc('figure', figsize=(20, 12))
i = 0
for site in data.columns:
    i += 1
    print i
    plot_ecm(data,'20150901', '20150930', site, result, df_info.ix[site].Threshold, df_info.ix[site].DaytimeAve)
    if site in r2.index:
        plt.savefig("./fig/ecm/"+site+".jpeg", dpi = 150)
    else:
        plt.savefig("./fig/nonecm/"+site+".jpeg", dpi = 100)
    plt.close()


#==============================================================================
#calculate the daytime average， removed all duplicates
#arr = ['1-50', '51-100', '101-200', '201-300', '301-400', '401-500', '501-600', '601-700']
#arr=['1-25','26-50']
#daytime_ave = []
#for f in arr:
#    data = pd.read_csv('./file/'+f+'.csv', sep=',')
#    data['time'] = pd.to_datetime(data['time'])
#    data.set_index('time', inplace=True)
#    
#    for site in data.columns:
#        df = DataFrame(data[site]).drop_duplicates()
#        ave = df.between_time('10:00','15:00').mean()
#        daytime_ave.append(ave)        
# 
#df_ave = DataFrame(pd.concat(daytime_ave))       
#df_ave.to_csv('./ave.csv')
#==============================================================================


#==============================================================================
# some trivial test, ignore it now
#  def average(arr):
#     return arr.sum() / arr.size 
#     
#     r2['duration'] = r2.end - r2.start
#     grouped = r2.groupby(level = 0)
#     g = grouped.agg({'saving' : ['sum', 'mean', 'size'], 'duration' : ['sum', average]})    
#==============================================================================


#deal with the original data
#i = 0
#chunks = pd.read_csv('./file/Aug.csv',sep=',',chunksize=400000,iterator=True)
#for chunk in chunks:    
#    i+=1    
#    filename='./file/'+str(i)+'.csv'    
#    chunk.to_csv(filename,index=False)   
#
#
#for i in range(1,10):
#    fin = './file/'+str(i)+'.csv' 
#    fout = './file/pivot'+str(i)+'.csv' 
#    data = pd.read_csv(fin, sep=',')
#    data.columns=['Name','Time','Value']
#    data = data.pivot(index='Time',columns='Name',values='Value')
#    data.to_csv(fout)
#
#dfs = []
#for i in range(1, 10):
#    fin = './file/pivot'+str(i)+'.csv'
#    data = pd.read_csv(fin, sep=',')
#    data.set_index('Time',inplace=True)
#    dfs.append(data)
#result = pd.concat(dfs, axis=1)
#result.to_csv('./file/AugPivot.csv')