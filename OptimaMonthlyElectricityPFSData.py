'''
      ======== This program is to combine "HHDATA" files exported from Optima ========

running...
'''
print(__doc__)

import os
import pandas as pd
import numpy as np

'''private functions'''
# this function is to check if data file has wrong dates selection, weekly
# args: int month, pd.DataFrame calendar, pd.DataFrame df
def confirmMonthlyDateRange (month,calendar,df):
    StartDate = pd.Timestamp(calendar.loc[month, 'Date'])
    EndDate = StartDate + pd.Timedelta(days = calendar.loc[month, 'Days']-1)
    print(StartDate, ' /', EndDate)
    return (df[df['READING DATE'] < StartDate].empty and df[df['READING DATE'] > EndDate].empty)

    # args: pd.DataFrame readings, List [] labels, int datapoints_expected, boolean prorate
def groupData (readings, labels, datapoints_expected, prorate):
    grouped = readings.groupby(labels)
    data_points = grouped['Daily Total'].agg([np.count_nonzero])
    data_missing = data_points[data_points['count_nonzero']!= datapoints_expected].copy()
    prorate_mask = data_missing['count_nonzero'].map(lambda x: datapoints_expected/x if x>0 else np.nan)
    prorate_mask.index = prorate_mask.index.map(lambda x: x[1])
    prorate_mask.dropna(inplace = True)
    if data_missing.empty:
        print('No data missing, please proceed')
        data_missing = pd.DataFrame()
    else:
        print('\n==== Below are the meters and sites with data missing for Month {0}, in total of {1} days===='.format(month, datapoints_expected))
        print(data_missing)
    site_grouped = readings.groupby(['SiteID'])
    reporting_consumption = site_grouped['Daily Total'].agg([np.sum])
    if prorate:
        for i, v in prorate_mask.iteritems():
            reporting_consumption.set_value(i,'sum', reporting_consumption.loc[i,'sum']*v)
    return reporting_consumption, data_missing

def combineFiles(files, key_string, dataheaders):
    data = [f for f in files if key_string in f]
    res = {}
    for i, item in enumerate(data):
        try:
            res[i] = pd.read_excel(item) [dataheaders]
        except:
            print("Check data source format, to match the custom headers noted in the script")      
    # Combine data and remove duplicates
    df = pd.concat(res, ignore_index = True)
    df = df.drop_duplicates() # This command is expensive
    return df
    
        
'''  Inputs '''
out_put = r'Electricity Data Reporting.xlsx'
month = 4
path = r'C:\Users\czhang0914\Desktop\Q2 milestones\Monthly Reporting Template Improvments\Optima Data Combiner'
key_string = "HHDATA"
config = "config info.xlsx"
dataheaders = ['OPTIMA Half Hourly DATA','READING DATE','CHANNEL TYPE','UNITS','Daily Total', 'Max Reading', 'Min Reading',
       'Data Source', 'Status', 'Site Name']
pfs_default = 273.97

''' Main Script'''
# Set configurations ==========================================================
os.chdir(path)
files = os.listdir(path)
hh = pd.read_excel(config, sheetname = 'HHmeters',converters = {'Meter':str})
pfs = pd.read_excel(config, sheetname = 'PFSmeters')
calendar = pd.read_excel(config, sheetname = '2017CalendarMonth', index_col = 'Month')
datapoints_expected = calendar.loc[month, 'Days']

# Combine data into a raw dataframe ==========================================
df = combineFiles(files, key_string, dataheaders)

# Check date range of the data and output results ============================
if not confirmMonthlyDateRange(month,calendar,df):
    print('Please check data source date range!!')
# Join meter and SiteID information
main_meter = pd.merge(df, hh, left_on = 'OPTIMA Half Hourly DATA', right_on = 'Meter', how = 'inner')
pfs_meter = pd.merge(df, pfs, left_on = 'OPTIMA Half Hourly DATA', right_on = 'PFSmeters', how = 'inner')

main_reporting, main_missing = groupData(main_meter, ['Meter','SiteID','Contract'], datapoints_expected, False)
pfs_reporting, pfs_missing = groupData(pfs_meter, ['PFSmeters', 'SiteID', 'Contract'], datapoints_expected, True)

# pfs Reporting ===================================================================

pfs_reporting = pd.merge(pfs, pfs_reporting, left_on = 'SiteID', right_index =True, how = 'outer')
pfs_reporting.loc[pfs_reporting['PFSmeters'] == 'no','sum'] = pfs_default = 273.97 * datapoints_expected

# Generate report
writer =pd.ExcelWriter(out_put)
main_missing.to_excel(writer, sheet_name ='dataMissing summary', startcol = 0)
pfs_missing.to_excel(writer, sheet_name ='dataMissing summary', startcol =6)
main_reporting.to_excel(writer, sheet_name ='mainMeter reporting')
pfs_reporting.to_excel(writer, sheet_name = 'pfsMeter reporting')
df.to_excel(writer, sheet_name ='organized data', index = False)
writer.save()
