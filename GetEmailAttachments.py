
"""
Created on Wed Jun 01 13:52:41 2016

@author: czhang0914
"""

import win32com.client
import os
import datetime as dt
import pandas as pd

LocalWorkDirectory='C:\\WinPython-64bit-2.7.10.3\\WinPython-64bit-2.7.10.3\\zcc\\TestMsg\\'
TargetEmailSubject='Lighting overide alarm'
TargetDate=pd.Timestamp(dt.datetime(2016,6,1,12,0,0))
TargetEndDate=pd.Timestamp(dt.datetime(2016,6,2,12,0,0))

# Progress bar, to count # of attachments collected
import sys
def ProgressBar(val,end_val,txt,bar_length=30):
    percent = float(val) / end_val
    hashes = '#' * int(round(percent * bar_length))
    spaces = ' ' * (bar_length - len(hashes))
    sys.stdout.write("\r {0} {1}: [{2}] {3}%".format(val,txt,hashes + spaces, int(round(percent * 100))))

# Collecte attachments
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case,
                                    # the inbox. You can change that number to reference
                                    # any other folder
messages = inbox.Items

print '\n'

for m in messages:
   if m.Subject.encode('ascii','ignore') == TargetEmailSubject:    
       att=m.Attachments
       
att_count=att.Count
count=0
for i in att:
    count+=1
    i.SaveAsFile(LocalWorkDirectory+str(count)+'.msg')
    ProgressBar(count,att_count,'Attachments downloaded')
    
# collect email.body
body_content=[]
sep='\n'
print '\n'

for i in range(count):    
    filelocation=LocalWorkDirectory+str(i+1)+'.msg'
    try:
        msg = outlook.OpenSharedItem(filelocation)
    except:
        pass
    text = msg.body.encode('ascii','ignore')
    body_content.append(text)
    ProgressBar(i+1,count,'Local Attachements Files Collected') 


# Remove empty lines, put every item into a list
output=sep.join(body_content)
output = os.linesep.join([s for s in output.splitlines() if s.strip()])
print '\n\n'+output

ls=output.split('\r\n')

# Organize data as Dataframe

df=pd.DataFrame(ls)
df=df.ix[:,0].str.split('->',expand=True)
tx=df.ix[:,0].str.split('   ',expand=True)
meter_index=tx.ix[:,0].str.split('/',expand=True)
status=tx.ix[:,1]
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
