import pyodbc
import xlsxwriter
import pandas as pd
import numpy as np


# Connect to SQL Server and set cursor
conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=LALUCKYSERVER\SQLEXPRESS;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()

# Get list of invoices
sqlstring = 'SELECT invc_no, invc_dat, cust_no, item_no, desc_1, desc_2, qty FROM dbo.IHSLIN00 WHERE DATEPART(MONTH, invc_dat) >=? AND sls_rep=?'


month = 9
sls_rep = 21


df = pd.read_sql(sqlstring, conn, params=(month, sls_rep))
df['item_no'] = df['item_no'].str.strip()
# df.columns = df.columns.str.strip()


# Load New years Itmes

fname = "Data/NewYear.xlsx"
newyear_df = pd.read_excel(fname, sheet_name='Sheet1')

columns = [ 'invc_no', 'invc_dat', 'cust_no', 'item_no', 'desc_1', 'desc_2', 'qty']
# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Out/New_Year_Purchase_History.xlsx', engine='xlsxwriter')

temp_df = pd.DataFrame(columns=columns)

for index, row in df.iterrows():
    #temp_list =[]
    #temp_df = pd.DataFrame(columns=columns)
    for idx2, item in newyear_df.iterrows():
        if row.item_no == item.item_no:
            print('Item', item)
            temp_df = temp_df.append(row)
            break
            #temp_list.append(row)

    '''
    if not temp_list:
        temp_df = pd.DataFrame(temp_list)
        sname = row.cust_no + str(row.invc_dat)
        temp_df.to_excel(writer, sheet_name=sname, index=False)
        temp_df.iloc[0:0]
        temp_list.clear()
    '''

temp_df.to_excel(writer, sheet_name='Sheet1', index=False)
writer.save()




