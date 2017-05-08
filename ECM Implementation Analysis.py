

import pandas as pd
import datetime as dt
import os

os.chdir(r'C:\Users\czhang0914\Downloads')

csvfilename='ECM_20170306_202105.csv'
targetfilename='ECM Implementation Jan.xlsx'
Today=pd.Timestamp(dt.datetime(2017,3,10,0,0,0))

#
s_day = pd.Timestamp(dt.datetime(2017,1,1,0,0,0))
e_day = pd.Timestamp(dt.datetime(2017,1,31,0,0,0))


fullEOP=pd.read_csv(csvfilename,sep=';',parse_dates=[11,12,13,19,34],dayfirst=True,encoding='utf-16')
cnames=list(fullEOP.columns.values)
#SiteContract=pd.read_excel('SiteContractList.xlsx',index_col=0)



statehistory=fullEOP['State History'].str.split('|',expand=True)

'''organize state history'''

for i in range(20):
   idx=1+i*4
   statehistory.iloc[:,idx]=statehistory.iloc[:,idx].apply(pd.to_datetime,dayfirst=True)

'''
 # Take first and second state history records
'''
sh=statehistory.iloc[:,[0,1,4,5]]
#sh.iloc[:,[1,5]]=sh.iloc[:,[1,5]].apply(pd.to_datetime,dayfirst=True)
sh.columns=['StateHistory First Record','StateHistory First Record Date','StateHistory Second Record','StateHistory Second Record Date']

stats=fullEOP.iloc[:,[0,1,5,6,10,11,12,15,16,17,7]].copy()
stats=pd.concat([stats,sh],axis=1)
description = fullEOP.iloc[:,[0,1,2,3,4,5,6,7,8,14,15]].copy()
#stats.ix[stats['Implementation Date'].isnull(),'Implementation Date']=stats['StateHistory First Record Date']


''' insert new columns to compute in-year remaining and in-year existing savings
input today'''

EndofYear=pd.Timestamp(dt.datetime(2018,2,5,0,0,0)) # First day of first week in 2018
StartofYear=pd.Timestamp(dt.datetime(2017,1,30,0,0,0))
Dyear=EndofYear-StartofYear
Rdays=EndofYear-Today
Rd=Rdays/Dyear
#stats['in-year remaining days']=stats.loc[stats['State'].isin(['Realized','Implemented']),'StateHistory First Record Date'].map(lambda x:EndofYear-x if x>=StartofYear else None)
#stats['in-year remaining savings GBP']=stats['in-year remaining days']/Dyear*stats['Estimated Cost £']
#stats['in-year from Today KWH']=Rd*stats['Estimated Consumption kWh']
#stats['in-year from Today GBP']=Rd*stats['Estimated Cost £']
#stats['in-year ECM effective days']=stats['StateHistory First Record Date'].map(lambda x:Today-x if x>StartofYear and Today>x else None)
#stats['in-year ECM effective GBP']=stats['in-year ECM effective days']/Dyear*stats['Estimated Cost £']

'''remove ECM in the state 'Internal Validation'''
stats=stats[stats.iloc[:,4]!='Internal Validation']
mask = (stats['StateHistory First Record Date'] >= s_day) & (stats['StateHistory First Record Date'] <= e_day)

implemented_ECMs = stats[stats.iloc[:,11].isin(['Implemented'])].copy()
implemented_ECMs = implemented_ECMs[mask]

realized_ECMs=stats[stats.iloc[:,11].isin(['Realized'])].copy()
realized_ECMs = realized_ECMs[mask]
#ECM =stats.iloc[:,[0,1,2,3,4,8,9,10,11,12]]
implemented=implemented_ECMs.iloc[:,[0,1,2,3,4,8,9,10,11,12]]
realized=realized_ECMs.iloc[:,[0,1,2,3,4,8,9,10,11,12]]
'''in period effective savings'''
implemented['effective days'] = implemented['StateHistory First Record Date'].map(lambda x: e_day -x)
implemented['effective GBP'] = implemented['effective days']/Dyear*implemented['Estimated Cost £']
implemented['effective kWh'] = implemented['effective days']/Dyear*implemented['Estimated Consumption kWh']

realized['effective days'] = realized['StateHistory First Record Date'].map(lambda x: e_day -x)
realized['effective GBP'] = realized['effective days']/Dyear*realized['Estimated Cost £']
realized['effective kWh'] = realized['effective days']/Dyear*realized['Estimated Consumption kWh']
#ECM_BMSc=ECM.loc[(ECM['ECM Category']=='BMS')&(ECM['State'].isin(['New','To be implemented','Updated','Investigate'])),['Code','Site ID']].copy()
#ECM_BMScount=len(ECM_BMSc['Site ID'].value_counts())
#ECM_BMS_rest=ECM.loc[(ECM['ECM Category']=='BMS')&(ECM['State'].isin(['New','To be implemented','Updated','Investigate'])&(ECM['Type Code']!='S046')&(ECM['Type Code']!='S047')),['Code','Site ID','State','Type Code']].copy()
#ECM_BMS_rest_count=ECM_BMS_rest.groupby(['State']).agg({'Code':lambda x:len(x),'Site ID':lambda x:len(x.value_counts())})
#
#All=ECM.groupby(['ECM Category','Type Code','Energy','State'],as_index=False).agg({'Code':lambda x:len(x),
#                                                            'Site ID':lambda x:len(x.value_counts()),
#                                                            'Estimated Cost £':lambda x:sum(x),
#                                                            'Estimated Consumption kWh':lambda x:sum(x)})
#All=All.rename(columns={'Code':'Number of ECMs','Site ID':'Number of Sites'})
#All=All[['ECM Category','Type Code','Energy','State','Number of ECMs','Number of Sites','Estimated Consumption kWh','Estimated Cost £']]

Implemented=implemented.groupby(['ECM Category','Type Code','Energy','State'],as_index=False).agg({'Code':lambda x:len(x),
                                                            'Site ID':lambda x:len(x.value_counts()),
                                                            'Estimated Cost £':lambda x:sum(x),
                                                            'Estimated Consumption kWh':lambda x:sum(x),
                                                            'effective GBP':lambda x:sum(x),
                                                            'effective kWh':lambda x:sum(x)})
Implemented=Implemented.rename(columns={'Code':'Number of ECMs','Site ID':'Number of Sites'})
Implemented['Average kWh'] = Implemented['Estimated Consumption kWh']/Implemented['Number of ECMs']
Implemented['Average Cost £'] = Implemented['Estimated Cost £']/Implemented['Number of ECMs']
Implemented=Implemented[['ECM Category','Type Code','Energy','State','Number of ECMs','Number of Sites','Estimated Consumption kWh','Estimated Cost £','Average kWh','Average Cost £','effective kWh','effective GBP']]


implemented_sitewise = implemented.groupby(['Site ID','Energy','State'],as_index=False).agg({'Code':lambda x:len(x),
                                                            'Estimated Cost £':lambda x:sum(x),
                                                            'Estimated Consumption kWh':lambda x:sum(x),
                                                            'effective GBP':lambda x:sum(x),
                                                            'effective kWh':lambda x:sum(x)})
implemented_sitewise=implemented_sitewise.rename(columns={'Code':'Number of ECMs'})

Realized=realized.groupby(['ECM Category','Type Code','Energy','State'],as_index=False).agg({'Code':lambda x:len(x),
                                                            'Site ID':lambda x:len(x.value_counts()),
                                                            'Estimated Cost £':lambda x:sum(x),
                                                            'Estimated Consumption kWh':lambda x:sum(x),
                                                            'effective GBP':lambda x:sum(x),
                                                            'effective kWh':lambda x:sum(x)})
Realized=Realized.rename(columns={'Code':'Number of ECMs','Site ID':'Number of Sites'})
Realized['Average kWh'] = Realized['Estimated Consumption kWh']/Realized['Number of ECMs']
Realized['Average Cost £'] = Realized['Estimated Cost £']/Realized['Number of ECMs']
Realized=Realized[['ECM Category','Type Code','Energy','State','Number of ECMs','Number of Sites','Estimated Consumption kWh','Estimated Cost £','Average kWh','Average Cost £','effective kWh','effective GBP']]


realized_sitewise = realized.groupby(['Site ID','Energy','State'],as_index=False).agg({'Code':lambda x:len(x),
                                                            'Estimated Cost £':lambda x:sum(x),
                                                            'Estimated Consumption kWh':lambda x:sum(x),
                                                            'effective GBP':lambda x:sum(x),
                                                            'effective kWh':lambda x:sum(x)})
realized_sitewise = realized_sitewise.rename(columns={'Code':'Number of ECMs'})

#TCC=Implemented.iloc[:,[0,1]].copy()
#TCC=TCC.drop_duplicates()
#TypeCodes=TCC.groupby('ECM Category')['Type Code'].apply(lambda x: x.tolist())
#TypeCodes=TypeCodes.to_frame()


# pivot table for ECM summary
#YTF=pd.pivot_table(ECM,values=['Estimated Cost  \xc2\xa3 ','Estimated Consumption'],index=['ECM Category'],columns=['State'],aggfunc=np.sum,margins=True)
'''
# list state history periods time
statehistory.iloc[:,3]=statehistory.iloc[:,1].map(lambda x:Today-x)
for i in range(1,20):
    idx=i*4+1
    try:    
        statehistory.iloc[:,idx+2]=statehistory.iloc[:,idx-4]-statehistory.iloc[:,idx]
    except:
        pass
    '''

with pd.ExcelWriter(targetfilename) as writer:
#    TypeCodes.to_excel(writer,sheet_name='Implemented_TypeCodes')
#    All.to_excel(writer,sheet_name='All',index=False) 
    Implemented.to_excel(writer,sheet_name='Implemented',index=False) 
    Realized.to_excel(writer,sheet_name='Realized',index=False)                     
    implemented_sitewise.to_excel(writer,sheet_name='implemented_sitewise',index=False)
    realized_sitewise.to_excel(writer,sheet_name = 'realized_sitewise',index = False)
    