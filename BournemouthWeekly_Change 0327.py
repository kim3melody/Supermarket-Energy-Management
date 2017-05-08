

import pandas as pd
import numpy as np
import datetime
import os


# progress bar
import sys

def ProgressBar(val,end_val, bar_length=11):
    percent = float(val) / end_val
    hashes = '#' * int(round(percent * bar_length))
    spaces = ' ' * (bar_length - len(hashes))
    sys.stdout.write("\rProgress: [{0}] {1}%".format(hashes + spaces, int(round(percent * 100))))
    sys.stdout.flush()

# Main Script
path=r'C:\Users\czhang0914\Desktop\KG\Python Test Week 25'
os.chdir(path)
files = os.listdir(path)

# import data and  pre-process
dd=[]
for i in files:
    d=pd.read_csv(i,index_col=0)
    d.index=d.index.map(lambda x: x+' 00:00:00' if len(x)<12 else x+':00' if len(x)<17 else x)
    d.index=pd.to_datetime(d.index,format='%d/%m/%Y %H:%M:%S')
    d = d.reset_index().drop_duplicates(subset='index', keep='last').set_index('index')
    dd.append(d)

    
#wk = i[-6:-4]
wk = '35'
data=pd.concat(dd,axis=1)
data=data.fillna(0)

# import csv file
# cal=pd.read_csv('adlams8.csv')
# TargetFile=r'test.xlsx'
# Configure pump groups
site={'adlams':[
                 [['ADLM_PLC_FLOW'],['ADLM_PLC2_PUMP1_POWER','ADLM_PLC2_PUMP2_POWER','ADLM_PLC2_PUMP3_POWER']]
                 ],
      'ald llp':[
                  [['ALLL_R5_6_FLOW'],['ALLL_LLP1_POWER','ALLL_LLP2_POWER','ALLL_LLP3_POWER']]
                  ],
      'Ald pp':[
                 [['ALHL_R7_8_FLOW'],['ALHL_PP1_POWER','ALHL_PP2_POWER','ALHL_PP3_POWER','ALHL_PP4_POWER']]
                 ],
      'BB':[
             [['ALHL_BOUR_FLOW'],['ALHL_BB1_POWER','ALHL_BB2_POWER','ALHL_BB3_POWER']]
             ],
      'km':[
             [['KMHL1_FT0213'],['KMHL1_CC1_POWER','KMHL1_CC2_POWER','KMHL1_CC3_POWER']],
             [['KMHL1_FT0215'],['KMHL1_NM1_POWER','KMHL1_NM2_POWER','KMHL1_NM3_POWER']],
             [['KMHL1_FT0217'],['KMHL1_HT1_POWER','KMHL1_HT2_POWER','KMHL1_HT3_POWER']]
             ],
      'Longham':[
                   [['LOHLPH_HLPOF_FLOW'],['LOHLPH_POWER_1','LOHLPH_POWER_2']],
                   [['LOHLPH_HLPOF_FLOW'],['LOHLPH_HLP1_POWER','LOHLPH_HLP2_POWER','LOHLPH_HLP3_POWER','LOHLPH_HLP4_POWER']]                   
                   ],
      'Matchams':[
                    [['MAPLC2_RAW_FLOW'],['MAABS_VS1_POWER','MAABS_VS2_POWER','MAABS_VS3_POWER']]     
                    ],
      'sbm':[
               [['SBM_CH_SCADA_FLOW'],['SBM_IP_BP1_POWER','SBM_IP_BP2_POWER','SBM_IP_BP3_POWER','SBM_IP_BP4_POWER']]
               ],
      'sway':[
               [['KMHL1_FT0220'],['KMHL1_SW1_POWER','KMHL1_SW2_POWER','KMHL1_SW3_POWER']]
               ],
      'twp':[
              [['KMTW_FAW_FLW'],['KMTW_P1_PWR','KMTW_P2_PWR','KMTW_P3_PWR']]
              ],
      'wg':[
              [['WG_HALE-1_BOREHOLE_FLOW','WG_HALE-1A_BOREHOLE_FLOW','WG_HALE-2_BOREHOLE_FLOW','WG_HALE-3_BOREHOLE_FLOW'],['WG_HALE-1_SUPPLY']],
              [['WG_HALE-4_BOREHOLE_FLOW'],['WG_HALE-4_SUPPLY_POWER']]
              ]
      }

# All tariffs have the same time pieces [0-7,7-9,9-16,16-19,19-20,20-0]
tariff={
        'Ald pp':      [0.07132,0.08833,0.09188,0.15699,0.09188,0.08833],
        'sbm':         [0.07404,0.09203,0.10082,0.19116,0.10082,0.09203],
        'ald llp':     [0.07132,0.08833,0.09188,0.15699,0.09188,0.08833],
        'Matchams':    [0.07432,0.09218,0.10097,0.19131,0.10097,0.09218],
        'km':          [0.07147,0.08842,0.09197,0.15708,0.09197,0.08842],
        'wg':          [0.07162,0.08687,0.09042,0.15553,0.09042,0.08687],
        'Longham':     [0.07132,0.08833,0.09188,0.15699,0.09188,0.08833],
        'BB':          [0.07132,0.08833,0.09188,0.15699,0.09188,0.08833],
        'adlams':      [0.07141,0.08842,0.09197,0.15708,0.09197,0.08842],
        'sway':        [0.07147,0.08842,0.09197,0.15708,0.09197,0.08842],
        'twp':         [0.07147,0.08842,0.09197,0.15708,0.09197,0.08842]      
        }

res={}
weekly_summary={}
daily_summary={}

DateTime=data.index
Date=DateTime.map(lambda x: x.date)
Time=DateTime.map(lambda x: x.time)
Week=DateTime.map(lambda x: str(x.year)+'-'+str((x+datetime.timedelta(days=1)).isocalendar()[1]))
Day=DateTime.map(lambda x: x.dayofweek)
Day_Night=DateTime.map(lambda x: 'night' if x.hour<7 else 'day')

# Loop through the site list, and pump groups
for i in site:
    temp=pd.DataFrame(index=DateTime)
    temp['Date']=Date;temp['Time']=Time;temp['Week']=Week; temp['day/night']=Day_Night; temp['Day']=Day 
    n=0 # Pump Group index
    weekly_summary[i]={}
    daily_summary[i]={}
    for pg in site[i]: 
        # tariff rules
        t=tariff[i]
        temp['New tariff']=DateTime.map(lambda x:  t[0] if x.dayofweek>=5 and x.hour<7                   # Sat, Sun before 7   => t[0]
                                                           else  t[1] if x.dayofweek>=5                  # Sat, Sun after 7    => t[1]
                                                           else  t[0] if x.hour<7                        # Mon-Fri before 7    => t[0]
                                                           else  t[1] if x.hour<9 or x.hour>19           # Mon-Fri 7-9, 20-24  => t[1]
                                                           else  t[2] if x.hour<16 or x.hour >18         # Mon-Fri 9-16, 19-20 => t[2]
                                                           else  t[3])                                   # Mon-Fri 16-19       => t[3]

        n+=1
        sn=str(n)
        pi=0 # pump index
        fi=0 # flow index
        MLidx='#ML-'+sn; POWERidx='#POWER-'+sn; Operation='Operation '+sn; type='type '+sn
        temp[MLidx]=0; temp[POWERidx]=0; temp[Operation]=0; temp[type]=''
        # Flow meters
        for f in pg[0]:
            fi+=1
            temp[f]=data[f].copy()
            temp[MLidx]+=temp[f]
        # Power meters
        for p in pg[1]: 
            pi+=1
            temp[p]=data[p].copy()
            temp[POWERidx] += temp[p]
            temp[Operation] += temp[p].map(lambda x: 1 if x>12 else 0) 
            temp[type] += temp[p].map(lambda x: str(pi)+'/' if x>12 else '-/')
        # ML, KWH and Cost for this pump group
        temp['ML '+sn]=temp[MLidx]*5/60/24
        temp['ML '+sn]=temp['ML '+sn].fillna(0)
        temp['KWH '+sn]=temp[POWERidx]*5/60 
        temp['KWH/ML '+sn]=0
        temp['KWH/ML '+sn]=np.where(temp['ML '+sn]>0,temp['KWH '+sn]/ temp['ML '+sn],temp['KWH/ML '+sn])
#       temp['KWH/ML '+sn]=(temp['KWH '+sn]/ temp['ML '+sn]).where(temp['ML '+sn]>0)
        temp['Cost '+sn]=temp['KWH '+sn]*temp['New tariff']

        daily_summary[i][n]=temp.groupby(['Date']).agg({'Cost '+sn:lambda x:sum(x),'KWH '+sn:lambda x:sum(x),'ML '+sn:lambda x: sum(x)})
        daily_summary[i][n] = daily_summary[i][n][['ML '+sn,'Cost '+sn,'KWH '+sn]]
        weekly_summary[i][n]=temp.groupby(['Week']).agg({'Cost '+sn:lambda x:sum(x),'KWH '+sn:lambda x:sum(x),'ML '+sn:lambda x: sum(x)})
        weekly_summary[i][n] = weekly_summary[i][n][['ML '+sn,'Cost '+sn,'KWH '+sn]]
    res[i]=temp
    
#       ==== Special Case wg =====
wg = res['wg']
wg = wg.drop(['type 1','type 2'],axis=1)
p1=wg['WG_HALE-1_BOREHOLE_FLOW'].map(lambda x: '1/' if x>1 else '-/')
p1A=wg['WG_HALE-1A_BOREHOLE_FLOW'].map(lambda x: '1A/' if x>1 else '-/')
p2=wg['WG_HALE-2_BOREHOLE_FLOW'].map(lambda x: '2/' if x>1 else '-/')
p3=wg['WG_HALE-3_BOREHOLE_FLOW'].map(lambda x: '3/' if x>1 else '-/')
p4=wg['WG_HALE-4_BOREHOLE_FLOW'].map(lambda x: '4/' if x>1 else '-/')

p1o=wg['WG_HALE-1_BOREHOLE_FLOW'].map(lambda x: 1 if x>1 else 0)
p1Ao=wg['WG_HALE-1A_BOREHOLE_FLOW'].map(lambda x: 1 if x>1 else 0)
p2o=wg['WG_HALE-2_BOREHOLE_FLOW'].map(lambda x: 1 if x>1 else 0)
p3o=wg['WG_HALE-3_BOREHOLE_FLOW'].map(lambda x: 1 if x>1 else 0)
p4o=wg['WG_HALE-4_BOREHOLE_FLOW'].map(lambda x: 1 if x>1 else 0)

wg['Operation'] = p1o+p1Ao+p2o+p3o+p4o
wg['type']=p1+p1A+p2+p3+p4
wg['Tariff for pump 4']=0.10182
wg=wg.rename(columns={'Cost 2':'Cost 4','ML 2':'ML 4','KWH 2':'KWH 4','#ML-2':'#ML-4','KWH/ML 2':'KWH/ML 4'})
wg['ML'] = wg['ML 1']+wg['ML 4']
wg.drop(['Operation 1','Operation 2'],axis =1 , inplace =True)
wg['Cost 4']=wg['KWH 4']*wg['Tariff for pump 4']
wg['Total cost']=wg['Cost 1']+wg['Cost 4']
res['wg']=wg
daily_summary['wg'][2]=wg.groupby(['Date']).agg({'Total cost': lambda x: sum(x)})
#daily_summary['wg'][2]=wg.groupby(['Date']).agg({'ML 4': lambda x:sum(x),'Cost 4': lambda x:sum(x)})
daily_summary['wg'][1]=wg.groupby(['Date']).agg({'ML': lambda x:sum(x)})
weekly_summary['wg'][2]=wg.groupby(['Week']).agg({'Total cost': lambda x: sum(x)})
#weekly_summary['wg'][2]=wg.groupby(['Week']).agg({'ML 4': lambda x:sum(x),'Cost 4': lambda x:sum(x)})
weekly_summary['wg'][1]=wg.groupby(['Week']).agg({'ML': lambda x:sum(x)})

#     ===== site km rename =======
res['km'] = res['km'].rename(columns={'ML 1':'ML CC','Cost 1':'Cost CC','KWH/ML 1':'KWH/ML CC','KWH 1':'KWH CC',
                                      'ML 2':'ML NM','Cost 2':'Cost NM','KWH/ML 2':'KWH/ML NM','KWH 2':'KWH NM',
                                      'ML 3':'ML HT','Cost 3':'Cost HT','KWH/ML 3':'KWH/ML HT','KWH 3':'KWH HT'
                                        })

daily_summary['km'][1] = daily_summary['km'][1][['ML 1','Cost 1','KWH 1']]
daily_summary['km'][2] = daily_summary['km'][2][['ML 2','Cost 2','KWH 2']]
daily_summary['km'][3] = daily_summary['km'][3][['ML 3','Cost 3','KWH 3']]

daily_summary['km'][1] = daily_summary['km'][1].rename(columns={'ML 1':'ML CC','Cost 1':'Cost CC','KWH 1':'KWH CC'})
daily_summary['km'][2] = daily_summary['km'][2].rename(columns={'ML 2':'ML NM','Cost 2':'Cost NM','KWH 2':'KWH NM'})
daily_summary['km'][3] = daily_summary['km'][3].rename(columns={'ML 3':'ML HT','Cost 3':'Cost HT','KWH 3':'KWH HT'})

weekly_summary['km'][1] = weekly_summary['km'][1][['ML 1','Cost 1','KWH 1']]
weekly_summary['km'][2] = weekly_summary['km'][2][['ML 2','Cost 2','KWH 2']]
weekly_summary['km'][3] = weekly_summary['km'][3][['ML 3','Cost 3','KWH 3']]

weekly_summary['km'][1] = weekly_summary['km'][1].rename(columns={'ML 1':'ML CC','Cost 1':'Cost CC','KWH 1':'KWH CC'})
weekly_summary['km'][2] = weekly_summary['km'][2].rename(columns={'ML 2':'ML NM','Cost 2':'Cost NM','KWH 2':'KWH NM'})
weekly_summary['km'][3] = weekly_summary['km'][3].rename(columns={'ML 3':'ML HT','Cost 3':'Cost HT','KWH 3':'KWH HT'})

# ===========Longham format ===============
res['Longham'].drop(['Operation 1','type 1'],axis=1,inplace =True)
res['Longham'] = res['Longham'].rename(columns = {'':'',

                                                    })
# Output results to files
p=0
for i in site:
    p+=1
    with pd.ExcelWriter(i+wk+'.xlsx') as writer:
            res[i].to_excel(writer,sheet_name=i+'cal')
            pd.concat(daily_summary[i],axis=1).to_excel(writer,sheet_name=i+' daily summary')
            pd.concat(weekly_summary[i],axis=1).to_excel(writer,sheet_name=i+' weekly summary')
            ProgressBar(p,11)
            
     


