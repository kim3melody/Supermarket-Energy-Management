import pymssql
import csv

''' 
User credentials
'''
server = "192.168.70.116"
user = "sa"
password = "Elutions69"

'''
pymssql SQL query

Object types:
    conn: pymssql.Connection
    cursor: pymssql.Cursor
'''
sql_string = ' SELECT [Code] ,[EntityName],[Consumption],[Cost],[SiteName],[TypeCode],[EnergyName],[WFStateName], CONVERT(varchar(100), [IdentificationDate], 101) as "ID date",[ECMType],[Cost],[Descriptions],[Actions] FROM [Morrisons_v4_Sys].[dbo].[ECMs] where [Deleted] = 1 order by Code'

# sql_string = ' SELECT * FROM [Morrisons_v4_Sys].[dbo].[ECMs] where [Deleted] = 1 order by Code '


conn = pymssql.connect(server, user, password, "Morrisons_v4_Sys")
cursor = conn.cursor()
cursor.execute(sql_string)

count = 0
with open('deleted_ecms_0228.csv', 'wb') as fp:
    a = csv.writer(fp, delimiter=',')
    ssss = 0
    for row in cursor:
        count += 1
        ssss += 1
        data = []
        
        for i in range(0,13):
            if row[i]:
                if type( row[i] ) == type(1.2) or type( row[i] ) == type(1):
                    data.append( row[i] )
                else:
                    data.append( row[i].encode('utf-8').strip() )
            else:
                data.append( 'null' )
        print data
        a.writerow(data)


'''
Print number of rows in the query
Close connection
'''
print count
conn.close()