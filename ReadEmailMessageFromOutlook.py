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
TargetEndDate=pd.Timestamp(dt.datetime(2016,6,28,12,0,0))
# Progress bar, to count # of emails collected
import sys
def ProgressBar(count,txt,bar_length=30):
    hashes = '#' * count
    sys.stdout.write("\r {0} {1}".format(count,txt))
    
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
print('\n')

for m in messages:
   if m.Subject == TargetEmailSubject:     
       text = m.body
       body_content.append(text)
       count+=1
       ProgressBar(count,' emails collected')  
       
       
# Remove empty lines, put every item into a list
       
output=sep.join(body_content)
output = os.linesep.join([s for s in output.splitlines() if s.strip()])
output=output.replace('2/3','two thirds')
output=output.replace('O/R','OR')
print('\n\n'+output)
print(str(count)+' emails collected')
ls=output.split('\r\n')


# Organize data as Dataframe
df=pd.DataFrame(ls)
df=df.ix[:,0].str.split('->',expand=True)
tx=df.ix[:,0].str.split('   ',expand=True)
meter_index=tx.ix[:,0].str.split('/',expand=True)
status=tx.iloc[:,1]
# Parse date time
DateTime=pd.to_datetime(df.ix[:,1],format='%d %B %Y %H:%M:%S')

Summary=pd.concat([meter_index,status,DateTime],axis=1,ignore_index=True)
Summary.columns=['Site Name',' ','Overide label','Alarm Status','Time']
Summary['Site Name']=Summary['Site Name'].map(lambda x: x.rstrip(' EINC'))
Result=Summary.sort_values(by=['Site Name','Overide label','Time'],ascending=[1,1,1])
Result.to_excel('BMS Lighting Overide Alarm_FULL.xlsx',index=False)

# Filter status and date
# Filter status and date
Result_filtered=Result[(Result['Alarm Status'].str.contains('OCCURRED'))&(Result['Time']>=TargetDate)&(Result['Time']<TargetEndDate)]
Result_filtered.to_excel('BMS Lighting Overide Alarm_OCCURRED.xlsx',index=False)
