import pyodbc
import xlrd
import pandas as pd
import numpy as np


file_name = "Out/address_phone_combined_f_final.xlsx"
df = pd.read_excel(file_name, sheet_name='Sheet1')
# Remove all NaN and replace with blank space
df = df.fillna('')

conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=MULTIMEDIAPC\SQLEXPRESS;DATABASE=pbsdataDEMO;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()

'''
sql = "SELECT top 1000 * FROM dbo.CUSFIL00"
cur.execute(sql)

for row in cur:
    print('row = %r' % (row,))

'''
rem_addr2 = "UPDATE dbo.CUSFIL00 SET addr_2=null WHERE cust_no=?"
for index, row in df.iterrows():
    try:
        cur.execute(rem_addr2, row.cust_no)
    except Exception as e:
        print(e)

sqlstring = "UPDATE dbo.CUSFIL00 SET city=?,state=?,zipcode=?,addr_3=? WHERE cust_no=?"
#sqlstring = "UPDATE dbo.CUSFIL00 SET addr_2=? WHERE cust_no=?"

for index, row in df.iterrows():
    values = [row.city, row.state, row.zip, row.addr_3, row.cust_no]
    try:
        cur.execute(sqlstring, values)
    except Exception as e:
        print(e)

#results = cur.fetchall()

cur.close()
conn.commit()

'''
with open('Out/update_log.txt', 'w') as f:
    for row in results:
        f.write("%s\n" % str(row))
'''