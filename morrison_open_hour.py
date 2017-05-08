# -*- coding: utf-8 -*-
"""
Created on Wed Aug 19 09:48:33 2015

@author: jzhu0922
"""

#get cafe open time
base_url = 'https://my.morrisons.com/storefinder/{0}'
import re
import urllib

def get_page(url):
    try:
        return urllib.request.urlopen(url).read()
    except:
        return ""
        
def get_open_time(page):
    #start_div = page.find('sf_cafe_div')
    start_div = page.find('sf_store_div')
    if start_div == -1:
        return ""
    start_dl = page.find('<dl class', start_div)
    end_dl = page.find('</dl>', start_dl + 1)
    time_str = page[start_dl:end_dl]
    return time_str
    
def get_store_name(page):
    start_div = page.find('Welcome to Morrisons')
    if start_div == -1:
        return ""
    end_div = page.find('</h3>',start_div)
    store_name = page[start_div : end_div]
    return store_name


f = open("store cafe_hour.csv", "w+")
for i in range(15):
    url = base_url.format(i)
    page = get_page(url)
    
    if page != "":
        store_name = get_store_name(page)
        time_str = get_open_time(page)
        time_str = re.sub(r'</?d[dlt]>',' ',time_str)
        
        s = ';'.join([str(i),store_name, time_str])+"\n"
        print(s)
        f.write(s.encode("UTF-8"))
f.close()


