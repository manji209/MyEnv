import pyodbc
import xlrd
import pandas as pd
import numpy as np


# Open excel sheet with order numbers and customer number
book = xlrd.open_workbook("Data/Orders_List_0904_0905.xlsx")
sheet = book.sheet_by_name("Table 1")

# Connect to SQL Server and set cursor
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
#conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()

columns = ['ord_no','item_no', 'qty_ord', 'cust_no']
df_total = pd.DataFrame(columns=columns)

def get_order(num, cust):
    sqlstring = 'SELECT ord_no, item_no, qty_ord FROM dbo.LINITM00 WHERE ord_no=?'
    df = pd.read_sql(sqlstring,conn,params={num})
    #Strip whitespaces from item_no
    df['item_no'] = df['item_no'].str.strip()
    df['cust_no'] = cust
    print(df)

    return df


for r in range(1, sheet.nrows):
    #print(int(sheet.cell(r,0).value))
    #print(sheet.cell(r,1).value)
    df_total = df_total.append(get_order(int(sheet.cell(r,0).value),sheet.cell(r,1).value))
    #df_total.append(get_order(110335, 'V032'))

#print(get_order(96,'H024'))

print(df_total)


# Write combined orders to excel
writer = pd.ExcelWriter('Out/special_orders_out.xlsx', engine='xlsxwriter')

df_total.to_excel(writer, sheet_name='Sheet1', index=False)

writer.save()

