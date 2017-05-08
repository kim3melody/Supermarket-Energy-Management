# -*- coding: utf-8 -*-
"""
Created on Fri Mar 31 15:31:23 2017

@author: czhang0914
"""

import pandas as pd
import os
import sys


''' config '''

path = r'C:\Users\czhang0914\Desktop\Q2 milestones\Monthly Reporting Template Improvments\Optima Monthly Invoice Parser'
index_string = '* START of DUoS DATA*'
file_string = 'SR130'
os.chdir(path)
files = os.listdir(path)
invo_files = [f for f in files if file_string in f and 'xlsx' in f.lower()]

dateColumns = ['Start Date', 'Reading Date', 'Invoice/Tax Point Date', 'Date Paid'] 
tariffColumns = ['Total CCL Charge £', 'Night Units £', 'Day Units £', 'Reactive Charge £', 'DUoS_Red Unit Charge £', 'DUoS_Amber Unit Charge £', 'DUoS_Green Unit Charge £']

DaysinMonth = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
Months = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']

''' private functions'''

def handleDuplicates(header):
    s = set()
    for i, item in header.iteritems():
        if item in s: 
            print('Duplicate headers handled: ' + item)
            header.set_value(i, item+'_duplicate')
        s.add(item)
        
def makeTable(table, header):
    table = table.iloc[1:]
    table.columns = header
    return table

dateParser = lambda x : pd.datetime.strptime(x, '%d/%m/%Y')


''' original table parsing, table concatenation, headers manipulating, date parsing'''

def parseInvoice(invo):
    # locate the second tariff table by looking for string 'Optima Energy Management System'
    for i, value in invo.loc[3].iteritems():
        if value == index_string:
            index_2 = i
        
    try:
        print('\n' + 'table_2 position: ' + str(index_2) + '\n')
    except:
        print("\n   =>   didn't find string: "+ index_string +"\n")
        sys.exit("error!")
    
    # Locate the month string 
    month_string = invo.iloc[1, 1]
    print('month_string: ' + month_string )
    for m in Months:
        if m in month_string.lower():
            month = m
    if not month:
        month = 'Unknow'
            
    table_1 = invo.iloc[3:, 0:index_2]
    table_2 = invo.iloc[3:, index_2+1:]
    
    header_1 = table_1.iloc[0]
    header_2 = table_2.iloc[0]
    
    # specify header_2 using prefix 'DUoS'
    for i, item in header_2.iteritems():
        item = 'DUoS_' + item
        header_2.set_value(i, item)
    
    handleDuplicates(header_1)
    handleDuplicates(header_2)
    
    table_1 = makeTable(table_1, header_1)
    table_2 = makeTable(table_2, header_2)
    table = pd.concat([table_1, table_2], axis = 1)
    
    for i in dateColumns:
        table[i] = table[i].map(dateParser)
    
    
    ''' filtering data '''
    # discard records with 0 consumption
    # discard records without DUoS Tariff
    table = table.loc[(table['Total Units (kWh)']!=0) & (pd.notnull(table['DUoS Tariff'])) ]
    table['Unit Rate'] = table[tariffColumns].sum(axis=1)/table['Total Units (kWh)']
    table.sort_values(by = ['Site Code', 'Total Units (kWh)'], ascending = [1,1], inplace = True)
    
    table.drop_duplicates(['Site Code'], 'last', inplace = True)
    return month, table

    
res = {}
for i in invo_files:
    month, res[i] = parseInvoice(pd.read_excel(i, header = None))