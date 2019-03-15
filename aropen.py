import pyodbc
import xlrd
import pandas as pd
import numpy as np
from datetime import datetime
from datetime import timedelta


# Connect to SQL Server and set cursor
#conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=DINHS-PC\SQLEXPRESS;DATABASE=pbsdataDEMO;UID=pbssqluser;PWD=Admin11')
conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=DINHPC\SQLEXPRESS;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()

sqlstring = "SELECT cust_no, doc_dat, due_dat, doc_typ, doc_no, appl_to_no, ref, amt_1, amt_2 FROM dbo.AROPEN00 WHERE cust_no=?"
customer_no = ""
temp_date = "03-07-2019"
cur_date = datetime.today() - timedelta(days=8)
date_30 = cur_date + timedelta(days=30)
date_60 = cur_date + timedelta(days=60)
current = 0
over_30 = 0
over_60 = 0
over_90 = 0

df = pd.read_sql(sqlstring,conn,params={customer_no})
df_inv = df.loc[df['doc_typ'] == 'I']
df_paid = df.loc[df['doc_typ'] == 'P']
df_credit = df.loc[df['doc_typ'] == 'C']

df_amount = df.pivot_table(values=['amt_1'], index=['doc_typ'], aggfunc={'amt_1':np.sum})

df_unpaid = pd.DataFrame(columns=['cust_no', 'doc_dat', 'due_dat', 'doc_typ', 'doc_no', 'appl_to_no', 'ref', 'amt_1', 'amt_2', 'net_due'])

for index, row in df.iterrows():
    if row.doc_typ == 'P' or row.doc_typ == 'C' or row.doc_typ == 'R':
        for index2, row2 in df.iterrows():
            if row.appl_to_no == row2.doc_no and row2.doc_typ == 'I':
                if row.amt_1 + row.amt_2 + row2.amt_1 == 0:
                    df.drop(index, inplace=True)
                    df.drop(index2, inplace=True)
                else:
                    df.loc[index2, 'amt_1'] = (row2.amt_1 + row.amt_1 + row.amt_2)
                    df.drop(index, inplace=True)
                    if row2.amt_1 == 0:
                        df.drop(index2, inplace=True)




'''

# Eliminate all Paid invoices
for index, row in df_paid.iterrows():
    for index2, row2 in df_inv.iterrows():
        if row.appl_to_no == row2.doc_no:
            if row.amt_1 + row.amt_2 + row2.amt_1 == 0:
                df_inv.drop(index2, inplace=True)
            else:
                #row2.amt_1 = (row2.amt_1 + row.amt_1 + row.amt_2)
                df_inv.loc[index2, 'amt_1'] = (row2.amt_1 + row.amt_1 + row.amt_2)
                #if df_inv.loc[index2, 'amt_1'] == 0:
                if row2.amt_1 == 0:
                    df_inv.drop(index2, inplace=True)

            #df_unpaid.append(row2)
            break

print(df_inv)


for index, row in df_credit.iterrows():
    for index2, row2 in df_inv.iterrows():
        if row.appl_to_no == row2.doc_no:
            if row.amt_1 + row.amt_2 + row2.amt_1 == 0:
                df_inv.drop(index2, inplace=True)
            else:
                #row2.amt_1 = (row2.amt_1 + row.amt_1 + row.amt_2)
                df_inv.loc[index2, 'amt_1'] = (row2.amt_1 + row.amt_1 + row.amt_2)
                #if df_inv.loc[index2, 'amt_1'] == 0:
                if row2.amt_1 == 0:
                    df_inv.drop(index2, inplace=True)
        # df_unpaid.append(row2)
            break

'''
#df_amount_net = df_inv.pivot_table(values=['amt_1'], index=['doc_typ'], aggfunc={'amt_1':np.sum})
df_amount_net = df.pivot_table(values=['amt_1', 'amt_2'], index=['doc_typ'], aggfunc={'amt_1':np.sum, 'amt_2':np.sum})

date_30 = cur_date + timedelta(days=30)
# Get aging report
for index, row in df.iterrows():
    if abs(cur_date.date() - row.doc_dat).days <= 30:
        current = current + row.amt_1 + row.amt_2
    elif abs(cur_date.date() - row.doc_dat).days >= 31 and abs(cur_date.date() - row.doc_dat).days <= 60:
        over_30 = over_30 + row.amt_1 + row2.amt_2
    elif abs(cur_date.date() - row.doc_dat).days >= 61 and abs(cur_date.date() - row.doc_dat).days <= 90:
        over_60 = over_60 + row.amt_1 + row2.amt_2
    else:
        over_90 = over_90 + row.amt_1 + row2.amt_2

print('Current Date: ', cur_date)
print('Current plus 30: ', date_30)
print('Current: ', current)
print('Past Due Over 30: ', over_30)
print('Past Due Over 60: ', over_60)
print('Past Due Over 90: ', over_90)


# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Out/aropen.xlsx', engine='xlsxwriter')

df_inv.to_excel(writer, sheet_name='Sheet1')
df.to_excel(writer, sheet_name='Sheet2')
writer.save()

#print(df)
#print(df_amount)
#print(df_inv)
#print(df_paid)
#print(df_inv)
#print(df_unpaid)
print(df_amount)
print(df_amount_net)





