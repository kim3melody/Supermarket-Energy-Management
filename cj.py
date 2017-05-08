# -*- coding: utf-8 -*-
"""
Created on Wed Jun 01 13:52:41 2016

@author: czhang0914
"""

import win32com.client
import os
import datetime as dt
import pandas as pd

LocalWorkDirectory='C:\\WinPython-64bit-2.7.10.3\\WinPython-64bit-2.7.10.3\\zcc\\TestMsg\\'
TargetEmailSubject='BMS Lighting overide alarm'
TargetDate=pd.Timestamp(dt.datetime(2016,6,27,12,0,0))
TargetEndDate=pd.Timestamp(dt.datetime(2016,7,7,12,0,0))
# Progress bar, to count # of emails collected
import sys
def ProgressBar(count,txt,bar_length=30):
    hashes = '#' * count
    sys.stdout.write("\r {0} {1}: {2}".format(count,txt,hashes))
    
# Collecte Emails from Outlook folder
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
exchangeaccount = outlook.Folders("MorrisonsLighting") # "6" refers to the index of a folder - in this case,
                                                       # the inbox. You can change that number to reference
                                                       # any other folder
inbox=exchangeaccount.Folders("Inbox")


messages = inbox.Items
body_content=[]
sep='\n'
count=0
print '\n'

for m in messages:
   if m.Subject.encode('ascii','ignore') == TargetEmailSubject:     
       text = m.body.encode('ascii','ignore')
       body_content.append(text)
       count+=1
       ProgressBar(count,' emails collected')  
       
       
# Remove empty lines, put every item into a list
       
output=sep.join(body_content)
output = os.linesep.join([s for s in output.splitlines() if s.strip()])
print '\n\n'+output

ls=output.split('\r\n')

# Organize data as Dataframe
df=pd.DataFrame(ls)
df=df.apply(lambda x: x.str.replace('2/3','two thirds'))
df=df.apply(lambda x: x.str.replace('O/R','OR'))
df=df.ix[:,0].str.split('->',expand=True)
tx=df.ix[:,0].str.split('   ',expand=True)
meter_index=tx.ix[:,0].str.split('/',expand=True)
site_name=meter_index.ix[:,0].apply(lambda x: x.rstrip(' EINC'))
overide_label= meter_index.ix[:,2]
import re
ECM_text=overide_label.apply(lambda x: re.sub('OR.*','',x))
ECM_text=ECM_text.apply(lambda x: re.sub('\\(.*','',x))
ECM_text=ECM_text.apply(lambda x:x[9:])
status=tx.ix[:,1]
# Parse date time
#try:
DateTime=pd.to_datetime(df.ix[:,1],format='%d %B %Y %H:%M:%S')
#except:
    #pass

Summary=pd.concat([site_name,overide_label,ECM_text,status,DateTime],axis=1,ignore_index=True)
Summary.columns=['Site Name','Overide Label','ECM Text','Alarm Status','Time']
#Summary['Site Name']=Summary['Site Name'].map(lambda x: x.rstrip(' EINC'))
Result=Summary.sort(['Site Name','ECM Text','Time'],ascending=[1,1,1])
Result.to_excel('BMS Lighting Overide Alarm_FULL.xlsx',index=False)

# Filter status and date
# Filter status and date
Result_filtered=Result[(Result['Alarm Status'].str.contains('OCCURRED'))&(Result['Time']>=TargetDate)&(Result['Time']<TargetEndDate)]
Result_filtered.to_excel('BMS Lighting Overide Alarm_OCCURRED.xlsx',index=False)
