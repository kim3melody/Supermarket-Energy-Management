

import pandas as pd
import datetime
import os

folder=r'C:\Users\czhang0914\Desktop\KG\Python Test Week 8'
os.chdir(folder)

# import csv file
#cal=pd.read_csv('adlams8.csv')
#TargetFile=r'test.xlsx'

# Configure pump groups

site={'adlams8':[
                 [['ADLM_PLC_FLOW'],['ADLM_PLC2_PUMP1_POWER','ADLM_PLC2_PUMP2_POWER','ADLM_PLC2_PUMP3_POWER']]
                 ],
      'ald llp8':[
                  [['ALLL_R5_6_FLOW'],['ALLL_LLP1_POWER','ALLL_LLP2_POWER','ALLL_LLP3_POWER']]
                  ],
      'Ald pp8':[
                 [['ALHL_R7_8_FLOW'],['ALHL_PP1_POWER','ALHL_PP2_POWER','ALHL_PP3_POWER','ALHL_PP4_POWER']]
                 ],
      'BB8':[
             [['ALHL_BOUR_FLOW'],['ALHL_BB1_POWER','ALHL_BB2_POWER','ALHL_BB3_POWER']]
             ],
      'km8':[
             [['KMHL1_FT0213'],['KMHL1_CC1_POWER','KMHL1_CC2_POWER','KMHL1_CC3_POWER']],
             [['KMHL1_FT0215'],['KMHL1_NM1_POWER','KMHL1_NM2_POWER','KMHL1_NM3_POWER']],
             [['KMHL1_FT0217'],['KMHL1_HT1_POWER','KMHL1_HT2_POWER','KMHL1_HT3_POWER']]
             ],
      'Longham 8':[
                   [['LOHLPH_HLPOF_FLOW'],['LOHLPH_HLP1_POWER','LOHLPH_HLP2_POWER','LOHLPH_HLP3_POWER','LOHLPH_HLP4_POWER']],
                   [['LOHLPH_HLPOF_FLOW'],['LOHLPH_POWER_1','LOHLPH_POWER_2']]
                   ],
      'Matchams 8':[
                    [['MAPLC2_RAW_FLOW'],['MAABS_VS1_POWER','MAABS_VS2_POWER','MAABS_VS3_POWER']]     
                    ],
      'sbm 8':[
               [['SBM_CH_SCADA_FLOW'],['SBM_IP_BP1_POWER','SBM_IP_BP2_POWER','SBM_IP_BP3_POWER','SBM_IP_BP4_POWER']]
               ],
      'sway8':[
               [['KMHL1_FT0220'],['KMHL1_SW1_POWER','KMHL1_SW2_POWER','KMHL1_SW3_POWER']]
               ],
      'twp8':[
              [['KMTW_FAW_FLW'],['KMTW_P1_PWR','KMTW_P2_PWR','KMTW_P3_PWR']]
              ],
      'wg 8':[
              [['WG_HALE-1_BOREHOLE_FLOW','WG_HALE-1A_BOREHOLE_FLOW','WG_HALE-2_BOREHOLE_FLOW','WG_HALE-3_BOREHOLE_FLOW','WG_HALE-4_BOREHOLE_FLOW'],['WG_HALE-1_BOREHOLE_FLOW']],
              [['WG_HALE-4_BOREHOLE_FLOW'],['WG_HALE-4_SUPPLY_POWER']]
              ]
      }

# import data
data1=pd.read_csv('Data1.csv')
data2=pd.read_csv('Data2.csv')
data3=pd.read_csv('Data3.csv')
data=pd.concat([data1,data2,data3],axis=1)
ti=data.iloc[:,0].map(lambda x: x+' 00:00:00' if len(x)<12 else x+':00' if len(x)<17 else x)
Date=pd.to_datetime(ti,format='%d/%m/%Y %H:%M:%S')

# Loop through the site list, and pump groups

for i in site:
    temp=pd.DataFrame(data=ti)
    n=0
    for pg in site[i]: 
        n+=1
        MLidx='#ML'+'-'+str(n)
        POWERidx='#POWER'+'-'+str(n)
        temp[MLidx]=0
        temp[POWERidx]=0
        for f in pg[0]: temp[f]=data[f].copy(); temp[MLidx]+=temp[f]
        for p in pg[1]: temp[p]=data[p].copy(); temp[POWERidx]+=temp[p]
        temp['KWH'+str(n)]=


    # parse datetime column
    cal['DateTime']=Date
    
    # calculating columns with forumulae
    cal['Power']=cal['ADLM_PLC2_PUMP1_POWER']+cal['ADLM_PLC2_PUMP2_POWER']+cal['ADLM_PLC2_PUMP3_POWER']
    cal['ML']=cal['ADLM_PLC_FLOW']*5/60/24
    cal['KWH']=cal['Power']*5/60
    
    # tariff rules
    cal['New tariff']=cal['DateTime'].map(lambda x:  0.07141 if x.dayofweek>=5 and x.hour<7
                                                       else  0.08842 if x.dayofweek>=5 and x.hour>=7
                                                       else  0.07141 if x.hour<7
                                                       else  0.08842 if x.hour<9 or x.hour>19
                                                       else  0.09197 if x.hour<16 or x.hour >18
                                                       else  0.15708)
    
    cal['Cost']=cal['KWH']*cal['New tariff']
    
    # Bourmouth team documents are using Sunday as the first day of a week
    cal['Week']=cal['DateTime'].map(lambda x: str(x.year)+'-'+str((x+datetime.timedelta(days=1)).isocalendar()[1]+1))
    
    
    pump1_status=cal['ADLM_PLC2_PUMP1_POWER'].map(lambda x: 1 if x>1 else 0) 
    pump2_status=cal['ADLM_PLC2_PUMP2_POWER'].map(lambda x: 1 if x>1 else 0)
    pump3_status=cal['ADLM_PLC2_PUMP3_POWER'].map(lambda x: 1 if x>1 else 0)
    
    cal['Operation']=pump1_status+pump2_status+pump3_status
                         
    cal['day/night']=cal['DateTime'].map(lambda x: 'night' if x.hour<7 else 'day')
    
    p1=cal['ADLM_PLC2_PUMP1_POWER'].map(lambda x: '1' if x>1 else '-')
    p2=cal['ADLM_PLC2_PUMP2_POWER'].map(lambda x: '2' if x>1 else '-')
    p3=cal['ADLM_PLC2_PUMP3_POWER'].map(lambda x: '3' if x>1 else '-')
    
    cal['pumps']= p1+'/'+p2+'/'+p3
    cal['day']=cal['DateTime'].map(lambda x: x.dayofweek)
    
    summary=cal.groupby(['Week']).agg({'ML':lambda x: sum(x),'Cost':lambda x:sum(x)})

with pd.ExcelWriter(TargetFile) as writer:
    cal.to_excel(writer,sheet_name='table',index=False)
    summary.to_excel(writer,sheet_name='summary',index=False)